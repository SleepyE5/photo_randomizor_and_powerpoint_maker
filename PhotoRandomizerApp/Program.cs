using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

var sourceFolder = args.Length > 0
    ? args[0]
    : Path.Combine(Directory.GetCurrentDirectory(), "PhotosInput");
var destinationFolder = args.Length > 1
    ? args[1]
    : Path.Combine(Directory.GetCurrentDirectory(), "PhotosOutput");

Directory.CreateDirectory(sourceFolder);
Directory.CreateDirectory(destinationFolder);

var photoExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp", ".heic", ".heif"
};

var videoExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".webm", ".m4v", ".mpeg", ".mpg", ".3gp"
};

var sourceFiles = Directory.EnumerateFiles(sourceFolder, "*.*", SearchOption.AllDirectories).ToList();
var ignoredVideoFiles = sourceFiles.Count(file => videoExtensions.Contains(Path.GetExtension(file)));
var photoFiles = sourceFiles
    .Where(file =>
    {
        var extension = Path.GetExtension(file);
        return photoExtensions.Contains(extension) && !videoExtensions.Contains(extension);
    })
    .ToList();

var uniquePhotoFiles = GetUniquePhotoFiles(photoFiles);

var copiedPhotos = new List<string>();

foreach (var file in uniquePhotoFiles)
{
    var sourceExtension = Path.GetExtension(file).ToLowerInvariant();
    var extension = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? ".jpg" : sourceExtension;
    var randomName = $"{Guid.NewGuid():N}{extension}";
    var outputPath = Path.Combine(destinationFolder, randomName);

    while (File.Exists(outputPath))
    {
        randomName = $"{Guid.NewGuid():N}{extension}";
        outputPath = Path.Combine(destinationFolder, randomName);
    }

    if (!TryCopyPhoto(file, sourceExtension, outputPath))
    {
        throw new IOException($"Failed to copy photo '{file}' to '{outputPath}'.");
    }

    NormalizeJpegOrientationInPlace(outputPath);
    copiedPhotos.Add(outputPath);
}

Console.WriteLine($"Done. Copied {uniquePhotoFiles.Count} photos to '{destinationFolder}'.");
var duplicatePhotoCount = photoFiles.Count - uniquePhotoFiles.Count;
if (duplicatePhotoCount > 0)
{
    Console.WriteLine($"Skipped {duplicatePhotoCount} duplicate photo(s) in '{sourceFolder}'.");
}
if (ignoredVideoFiles > 0)
{
    Console.WriteLine($"Ignored {ignoredVideoFiles} video file(s) in '{sourceFolder}'.");
}

var pptxPath = Path.Combine(destinationFolder, "Photos.pptx");
var slidesCreated = CreatePhotoPresentation(copiedPhotos, destinationFolder, pptxPath);

if (slidesCreated > 0)
{
    Console.WriteLine($"Created PowerPoint with {slidesCreated} slides: '{pptxPath}'");
}
else
{
    Console.WriteLine("No PowerPoint created because no PowerPoint-supported image formats were found.");
}

return;

static List<string> GetUniquePhotoFiles(IEnumerable<string> photoFiles)
{
    var uniquePhotoFiles = new List<string>();
    var seenHashes = new HashSet<string>(StringComparer.Ordinal);

    foreach (var photoFile in photoFiles)
    {
        var hash = ComputeFileHash(photoFile);
        if (seenHashes.Add(hash))
        {
            uniquePhotoFiles.Add(photoFile);
        }
    }

    return uniquePhotoFiles;
}

static string ComputeFileHash(string filePath)
{
    using var stream = File.OpenRead(filePath);
    var hashBytes = SHA256.HashData(stream);
    return Convert.ToHexString(hashBytes);
}

static bool TryCopyPhoto(string sourcePath, string sourceExtension, string outputPath)
{
    if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) && IsJpegLike(sourceExtension))
    {
        File.Copy(sourcePath, outputPath);
        return File.Exists(outputPath);
    }

    if (TryCreateNormalizedJpegOnMac(sourcePath, outputPath))
    {
        return true;
    }

    File.Copy(sourcePath, outputPath);
    return File.Exists(outputPath);
}

static int CreatePhotoPresentation(IEnumerable<string> imagePaths, string destinationFolder, string pptxPath)
{
    var powerPointImages = new List<(string FilePath, string Extension)>();
    var tempFolder = Path.Combine(destinationFolder, ".ppt-temp");

    foreach (var imagePath in imagePaths)
    {
        var extension = Path.GetExtension(imagePath).ToLowerInvariant();

        if (TryGetImagePartType(extension, out _))
        {
            powerPointImages.Add((imagePath, extension));
            continue;
        }

        if (IsHeicLike(extension) && TryConvertToJpegOnMac(imagePath, tempFolder, out var jpegPath))
        {
            TryAutoRotateOnMac(jpegPath);
            powerPointImages.Add((jpegPath, ".jpg"));
        }
    }

    if (powerPointImages.Count == 0)
    {
        return 0;
    }

    try
    {
        if (File.Exists(pptxPath))
        {
            File.Delete(pptxPath);
        }

        using (var presentationDocument = PresentationDocument.Create(pptxPath, PresentationDocumentType.Presentation))
        {
            var presentationPart = presentationDocument.AddPresentationPart();
            var slideLayoutPart = CreatePresentationStructure(presentationPart);
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("Presentation was not created.");
            var slideIdList = presentation.GetFirstChild<P.SlideIdList>()
                ?? throw new InvalidOperationException("SlideIdList was not created.");

            uint slideId = 256;
            var index = 1;

            foreach (var image in powerPointImages)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.AddPart(slideLayoutPart);
                var imagePartType = GetImagePartType(image.Extension);
                var imagePart = slidePart.AddImagePart(imagePartType, $"rIdImage{index}");

                using (var stream = File.OpenRead(image.FilePath))
                {
                    imagePart.FeedData(stream);
                }

                var (imageWidth, imageHeight) = GetImageSizeForSlide(image.FilePath);
                slidePart.Slide = CreateSlideWithImage($"rIdImage{index}", index, imageWidth, imageHeight);
                slidePart.Slide.Save();

                slideIdList.Append(new P.SlideId
                {
                    Id = slideId++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });

                index++;
            }

            presentationPart.Presentation.Save();
        }

        FixPptxContentTypes(pptxPath);
        return powerPointImages.Count;
    }
    finally
    {
        if (Directory.Exists(tempFolder))
        {
            Directory.Delete(tempFolder, recursive: true);
        }
    }
}

static P.Slide CreateSlideWithImage(string relationshipId, int index, long imageWidth, long imageHeight)
{
    const long slideWidth = 9_144_000L;
    const long slideHeight = 6_858_000L;
    var (targetWidth, targetHeight, offsetX, offsetY) = CalculateImagePlacement(
        imageWidth,
        imageHeight,
        slideWidth,
        slideHeight);

    var picture = new P.Picture(
        new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties { Id = 2U, Name = $"Photo {index}" },
            new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.BlipFill(
            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
            new A.Stretch(new A.FillRectangle())),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = offsetX, Y = offsetY },
                new A.Extents { Cx = targetWidth, Cy = targetHeight }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

    return new P.Slide(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                picture)),
        new P.ColorMapOverride(new A.MasterColorMapping()));
}

static (long Width, long Height) GetImageSizeForSlide(string imagePath)
{
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
    {
        return (4, 3);
    }

    var processStartInfo = new ProcessStartInfo
    {
        FileName = "sips",
        RedirectStandardOutput = true,
        RedirectStandardError = true
    };
    processStartInfo.ArgumentList.Add("-g");
    processStartInfo.ArgumentList.Add("pixelWidth");
    processStartInfo.ArgumentList.Add("-g");
    processStartInfo.ArgumentList.Add("pixelHeight");
    processStartInfo.ArgumentList.Add(imagePath);

    using var process = Process.Start(processStartInfo);
    if (process is null)
    {
        return (4, 3);
    }

    var output = process.StandardOutput.ReadToEnd();
    var errorOutput = process.StandardError.ReadToEnd();
    process.WaitForExit();
    if (process.ExitCode != 0)
    {
        return (4, 3);
    }

    long width = 0;
    long height = 0;
    var allOutput = output + "\n" + errorOutput;
    foreach (var line in allOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries))
    {
        if (line.Contains("pixelWidth:", StringComparison.OrdinalIgnoreCase))
        {
            width = ExtractFirstNumber(line);
        }
        else if (line.Contains("pixelHeight:", StringComparison.OrdinalIgnoreCase))
        {
            height = ExtractFirstNumber(line);
        }
    }

    return width > 0 && height > 0 ? (width, height) : (4, 3);
}

static (long TargetWidth, long TargetHeight, long OffsetX, long OffsetY) CalculateImagePlacement(
    long imageWidth,
    long imageHeight,
    long slideWidth,
    long slideHeight)
{
    var widthRatio = (double)slideWidth / imageWidth;
    var heightRatio = (double)slideHeight / imageHeight;
    var scale = Math.Min(widthRatio, heightRatio);

    var targetWidth = (long)Math.Round(imageWidth * scale);
    var targetHeight = (long)Math.Round(imageHeight * scale);
    var offsetX = (slideWidth - targetWidth) / 2;
    var offsetY = (slideHeight - targetHeight) / 2;

    return (targetWidth, targetHeight, offsetX, offsetY);
}

static long ExtractFirstNumber(string text)
{
    var match = Regex.Match(text, @"\d+");
    return match.Success && long.TryParse(match.Value, out var value) ? value : 0;
}

static SlideLayoutPart CreatePresentationStructure(PresentationPart presentationPart)
{
    var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
    var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
    var themePart = slideMasterPart.AddNewPart<ThemePart>();

    themePart.Theme = CreateTheme();
    slideLayoutPart.SlideLayout = CreateBlankSlideLayout();

    var slideLayoutRelId = slideMasterPart.GetIdOfPart(slideLayoutPart);
    slideMasterPart.SlideMaster = new P.SlideMaster(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()))),
        new P.ColorMap
        {
            Background1 = A.ColorSchemeIndexValues.Light1,
            Text1 = A.ColorSchemeIndexValues.Dark1,
            Background2 = A.ColorSchemeIndexValues.Light2,
            Text2 = A.ColorSchemeIndexValues.Dark2,
            Accent1 = A.ColorSchemeIndexValues.Accent1,
            Accent2 = A.ColorSchemeIndexValues.Accent2,
            Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4,
            Accent5 = A.ColorSchemeIndexValues.Accent5,
            Accent6 = A.ColorSchemeIndexValues.Accent6,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
        },
        new P.SlideLayoutIdList(
            new P.SlideLayoutId
            {
                Id = 2147483649U,
                RelationshipId = slideLayoutRelId
            }),
        new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
    slideMasterPart.SlideMaster.Save();

    var slideMasterRelId = presentationPart.GetIdOfPart(slideMasterPart);
    presentationPart.Presentation = new P.Presentation(
        new P.SlideMasterIdList(
            new P.SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = slideMasterRelId
            }),
        new P.SlideIdList(),
        new P.SlideSize { Cx = 9_144_000, Cy = 6_858_000, Type = P.SlideSizeValues.Screen4x3 },
        new P.NotesSize { Cx = 6_858_000, Cy = 9_144_000 },
        new P.DefaultTextStyle());

    return slideLayoutPart;
}

static P.SlideLayout CreateBlankSlideLayout()
{
    return new P.SlideLayout(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup()))),
        new P.ColorMapOverride(new A.MasterColorMapping()))
    {
        Type = P.SlideLayoutValues.Blank,
        Preserve = true
    };
}

static A.Theme CreateTheme()
{
    return new A.Theme(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" }))
            { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(
                    new A.LatinFont { Typeface = "Calibri" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(
                    new A.LatinFont { Typeface = "Calibri" },
                    new A.EastAsianFont { Typeface = string.Empty },
                    new A.ComplexScriptFont { Typeface = string.Empty }))
            { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.NoFill(),
                    new A.GroupFill()),
                new A.LineStyleList(
                    new A.Outline(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                    {
                        Width = 9525,
                        CapType = A.LineCapValues.Flat,
                        CompoundLineType = A.CompoundLineValues.Single,
                        Alignment = A.PenAlignmentValues.Center
                    }),
                new A.EffectStyleList(new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })))
            { Name = "Office" }))
    {
        Name = "Office Theme"
    };
}

static bool TryGetImagePartType(string extension, out PartTypeInfo imagePartType)
{
    switch (extension.ToLowerInvariant())
    {
        case ".jpg":
        case ".jpeg":
            imagePartType = ImagePartType.Jpeg;
            return true;
        case ".png":
            imagePartType = ImagePartType.Png;
            return true;
        case ".gif":
            imagePartType = ImagePartType.Gif;
            return true;
        case ".bmp":
            imagePartType = ImagePartType.Bmp;
            return true;
        case ".tif":
        case ".tiff":
            imagePartType = ImagePartType.Tiff;
            return true;
        default:
            imagePartType = default!;
            return false;
    }
}

static PartTypeInfo GetImagePartType(string extension)
{
    if (!TryGetImagePartType(extension, out var imagePartType))
    {
        throw new NotSupportedException($"Unsupported image extension for PowerPoint: {extension}");
    }

    return imagePartType;
}

static bool IsHeicLike(string extension)
{
    return extension is ".heic" or ".heif";
}

static bool TryConvertToJpegOnMac(string sourcePath, string tempFolder, out string jpegPath)
{
    jpegPath = string.Empty;
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
    {
        return false;
    }

    Directory.CreateDirectory(tempFolder);
    jpegPath = Path.Combine(tempFolder, $"{Guid.NewGuid():N}.jpg");

    var processStartInfo = new ProcessStartInfo
    {
        FileName = "sips",
        RedirectStandardOutput = true,
        RedirectStandardError = true
    };
    processStartInfo.ArgumentList.Add("-s");
    processStartInfo.ArgumentList.Add("format");
    processStartInfo.ArgumentList.Add("jpeg");
    processStartInfo.ArgumentList.Add(sourcePath);
    processStartInfo.ArgumentList.Add("--out");
    processStartInfo.ArgumentList.Add(jpegPath);

    using var process = Process.Start(processStartInfo);
    if (process is null)
    {
        return false;
    }

    process.WaitForExit();
    if (process.ExitCode != 0 || !File.Exists(jpegPath))
    {
        return false;
    }

    return true;
}

static bool TryCreateNormalizedJpegOnMac(string sourcePath, string outputPath)
{
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
    {
        return false;
    }

    Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");

    var processStartInfo = new ProcessStartInfo
    {
        FileName = "sips",
        RedirectStandardOutput = true,
        RedirectStandardError = true
    };
    processStartInfo.ArgumentList.Add("-s");
    processStartInfo.ArgumentList.Add("format");
    processStartInfo.ArgumentList.Add("jpeg");
    processStartInfo.ArgumentList.Add(sourcePath);
    processStartInfo.ArgumentList.Add("--out");
    processStartInfo.ArgumentList.Add(outputPath);

    using var process = Process.Start(processStartInfo);
    if (process is null)
    {
        return false;
    }

    process.WaitForExit();
    if (process.ExitCode != 0 || !File.Exists(outputPath))
    {
        return false;
    }

    return true;
}

static bool IsJpegLike(string extension)
{
    return extension is ".jpg" or ".jpeg";
}

static void NormalizeJpegOrientationInPlace(string imagePath)
{
    var extension = Path.GetExtension(imagePath).ToLowerInvariant();
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX) || !IsJpegLike(extension))
    {
        return;
    }

    if (!TryGetJpegExifOrientation(imagePath, out var orientation) || orientation == 1)
    {
        return;
    }

    switch (orientation)
    {
        case 2:
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 3:
            RunSipsInPlace(imagePath, "-r", "180");
            break;
        case 4:
            RunSipsInPlace(imagePath, "-f", "vertical");
            break;
        case 5:
            RunSipsInPlace(imagePath, "-r", "90");
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 6:
            RunSipsInPlace(imagePath, "-r", "90");
            break;
        case 7:
            RunSipsInPlace(imagePath, "-r", "270");
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 8:
            RunSipsInPlace(imagePath, "-r", "270");
            break;
    }

    SetJpegExifOrientation(imagePath, 1);
}

static bool TryGetJpegExifOrientation(string imagePath, out int orientation)
{
    orientation = 1;
    using var stream = File.OpenRead(imagePath);
    using var reader = new BinaryReader(stream);

    if (ReadUInt16BigEndian(reader) != 0xFFD8)
    {
        return false;
    }

    while (stream.Position < stream.Length)
    {
        if (reader.ReadByte() != 0xFF)
        {
            continue;
        }

        byte markerType;
        do
        {
            markerType = reader.ReadByte();
        } while (markerType == 0xFF);

        if (markerType == 0xD9 || markerType == 0xDA)
        {
            break;
        }

        var segmentLength = ReadUInt16BigEndian(reader);
        if (segmentLength < 2)
        {
            return false;
        }

        var segmentDataLength = segmentLength - 2;
        if (markerType != 0xE1)
        {
            stream.Position += segmentDataLength;
            continue;
        }

        var segmentData = reader.ReadBytes(segmentDataLength);
        if (segmentData.Length < 14)
        {
            continue;
        }

        if (segmentData[0] != 0x45 || segmentData[1] != 0x78 || segmentData[2] != 0x69 || segmentData[3] != 0x66)
        {
            continue;
        }

        var tiffStart = 6;
        bool littleEndian;
        if (segmentData[tiffStart] == 0x49 && segmentData[tiffStart + 1] == 0x49)
        {
            littleEndian = true;
        }
        else if (segmentData[tiffStart] == 0x4D && segmentData[tiffStart + 1] == 0x4D)
        {
            littleEndian = false;
        }
        else
        {
            return false;
        }

        var ifdOffset = ReadUInt32(segmentData, tiffStart + 4, littleEndian);
        var ifdPosition = tiffStart + (int)ifdOffset;
        if (ifdPosition + 2 > segmentData.Length)
        {
            return false;
        }

        var entryCount = ReadUInt16(segmentData, ifdPosition, littleEndian);
        for (var i = 0; i < entryCount; i++)
        {
            var entryOffset = ifdPosition + 2 + (i * 12);
            if (entryOffset + 12 > segmentData.Length)
            {
                return false;
            }

            var tag = ReadUInt16(segmentData, entryOffset, littleEndian);
            if (tag != 0x0112)
            {
                continue;
            }

            var type = ReadUInt16(segmentData, entryOffset + 2, littleEndian);
            var count = ReadUInt32(segmentData, entryOffset + 4, littleEndian);
            if (type != 3 || count < 1)
            {
                return false;
            }

            orientation = ReadUInt16(segmentData, entryOffset + 8, littleEndian);
            return true;
        }

        return false;
    }

    return false;
}

static ushort ReadUInt16BigEndian(BinaryReader reader)
{
    var bytes = reader.ReadBytes(2);
    return (ushort)((bytes[0] << 8) | bytes[1]);
}

static void SetJpegExifOrientation(string imagePath, ushort newOrientation)
{
    byte[] bytes;
    try
    {
        bytes = File.ReadAllBytes(imagePath);
    }
    catch
    {
        return;
    }

    if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
    {
        return;
    }

    var position = 2;
    while (position + 4 <= bytes.Length)
    {
        if (bytes[position] != 0xFF)
        {
            position++;
            continue;
        }

        var markerPosition = position;
        position++;
        while (position < bytes.Length && bytes[position] == 0xFF)
        {
            position++;
        }

        if (position >= bytes.Length)
        {
            return;
        }

        var markerType = bytes[position++];
        if (markerType == 0xD9 || markerType == 0xDA)
        {
            return;
        }

        if (position + 2 > bytes.Length)
        {
            return;
        }

        var segmentLength = (bytes[position] << 8) | bytes[position + 1];
        position += 2;
        if (segmentLength < 2)
        {
            return;
        }

        var segmentDataOffset = position;
        var segmentDataLength = segmentLength - 2;
        if (segmentDataOffset + segmentDataLength > bytes.Length)
        {
            return;
        }

        if (markerType == 0xE1 &&
            TrySetOrientationInExifSegment(bytes, segmentDataOffset, segmentDataLength, newOrientation))
        {
            try
            {
                File.WriteAllBytes(imagePath, bytes);
            }
            catch
            {
                // Keep original image if write fails.
            }

            return;
        }

        position = markerPosition + 2 + segmentLength;
    }
}

static bool TrySetOrientationInExifSegment(byte[] bytes, int offset, int length, ushort newOrientation)
{
    if (length < 14)
    {
        return false;
    }

    if (bytes[offset] != 0x45 || bytes[offset + 1] != 0x78 || bytes[offset + 2] != 0x69 || bytes[offset + 3] != 0x66)
    {
        return false;
    }

    var tiffStart = offset + 6;
    bool littleEndian;
    if (bytes[tiffStart] == 0x49 && bytes[tiffStart + 1] == 0x49)
    {
        littleEndian = true;
    }
    else if (bytes[tiffStart] == 0x4D && bytes[tiffStart + 1] == 0x4D)
    {
        littleEndian = false;
    }
    else
    {
        return false;
    }

    var ifdOffset = (int)ReadUInt32(bytes, tiffStart + 4, littleEndian);
    var ifdPosition = tiffStart + ifdOffset;
    if (ifdPosition + 2 > offset + length)
    {
        return false;
    }

    var entryCount = ReadUInt16(bytes, ifdPosition, littleEndian);
    for (var i = 0; i < entryCount; i++)
    {
        var entryOffset = ifdPosition + 2 + (i * 12);
        if (entryOffset + 12 > offset + length)
        {
            return false;
        }

        var tag = ReadUInt16(bytes, entryOffset, littleEndian);
        if (tag != 0x0112)
        {
            continue;
        }

        var type = ReadUInt16(bytes, entryOffset + 2, littleEndian);
        var count = ReadUInt32(bytes, entryOffset + 4, littleEndian);
        if (type != 3 || count < 1)
        {
            return false;
        }

        if (littleEndian)
        {
            bytes[entryOffset + 8] = (byte)(newOrientation & 0xFF);
            bytes[entryOffset + 9] = (byte)(newOrientation >> 8);
        }
        else
        {
            bytes[entryOffset + 8] = (byte)(newOrientation >> 8);
            bytes[entryOffset + 9] = (byte)(newOrientation & 0xFF);
        }

        bytes[entryOffset + 10] = 0;
        bytes[entryOffset + 11] = 0;
        return true;
    }

    return false;
}

static ushort ReadUInt16(byte[] data, int offset, bool littleEndian)
{
    if (littleEndian)
    {
        return (ushort)(data[offset] | (data[offset + 1] << 8));
    }

    return (ushort)((data[offset] << 8) | data[offset + 1]);
}

static uint ReadUInt32(byte[] data, int offset, bool littleEndian)
{
    if (littleEndian)
    {
        return (uint)(data[offset]
            | (data[offset + 1] << 8)
            | (data[offset + 2] << 16)
            | (data[offset + 3] << 24));
    }

    return (uint)((data[offset] << 24)
        | (data[offset + 1] << 16)
        | (data[offset + 2] << 8)
        | data[offset + 3]);
}

static void TryAutoRotateOnMac(string imagePath)
{
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
    {
        return;
    }

    var orientation = GetOrientationOnMac(imagePath);
    if (orientation is null || orientation == 1)
    {
        return;
    }

    switch (orientation)
    {
        case 2:
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 3:
            RunSipsInPlace(imagePath, "-r", "180");
            break;
        case 4:
            RunSipsInPlace(imagePath, "-f", "vertical");
            break;
        case 5:
            RunSipsInPlace(imagePath, "-r", "90");
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 6:
            RunSipsInPlace(imagePath, "-r", "90");
            break;
        case 7:
            RunSipsInPlace(imagePath, "-r", "270");
            RunSipsInPlace(imagePath, "-f", "horizontal");
            break;
        case 8:
            RunSipsInPlace(imagePath, "-r", "270");
            break;
    }

    RunSipsInPlace(imagePath, "-s", "orientation", "1");
}

static int? GetOrientationOnMac(string imagePath)
{
    var processStartInfo = new ProcessStartInfo
    {
        FileName = "sips",
        RedirectStandardOutput = true,
        RedirectStandardError = true
    };
    processStartInfo.ArgumentList.Add("-g");
    processStartInfo.ArgumentList.Add("orientation");
    processStartInfo.ArgumentList.Add(imagePath);

    using var process = Process.Start(processStartInfo);
    if (process is null)
    {
        return null;
    }

    var output = process.StandardOutput.ReadToEnd();
    process.WaitForExit();
    if (process.ExitCode != 0)
    {
        return null;
    }

    var line = output
        .Split('\n', StringSplitOptions.RemoveEmptyEntries)
        .FirstOrDefault(l => l.Contains("orientation:", StringComparison.OrdinalIgnoreCase));
    if (line is null)
    {
        return null;
    }

    var valueText = line[(line.IndexOf(':') + 1)..].Trim();
    return int.TryParse(valueText, out var value) ? value : null;
}

static void RunSipsInPlace(string imagePath, params string[] args)
{
    var processStartInfo = new ProcessStartInfo
    {
        FileName = "sips",
        RedirectStandardOutput = true,
        RedirectStandardError = true
    };

    foreach (var arg in args)
    {
        processStartInfo.ArgumentList.Add(arg);
    }
    processStartInfo.ArgumentList.Add(imagePath);

    using var process = Process.Start(processStartInfo);
    if (process is null)
    {
        return;
    }

    process.WaitForExit();
}

static void FixPptxContentTypes(string pptxPath)
{
    using var zip = ZipFile.Open(pptxPath, ZipArchiveMode.Update);
    var entry = zip.GetEntry("[Content_Types].xml");
    if (entry is null)
    {
        return;
    }

    XDocument document;
    using (var stream = entry.Open())
    {
        document = XDocument.Load(stream);
    }

    XNamespace contentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    var typesElement = document.Root;
    if (typesElement is null)
    {
        return;
    }

    var defaultXml = typesElement.Elements(contentTypesNs + "Default")
        .FirstOrDefault(e => string.Equals((string?)e.Attribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase));
    if (defaultXml is not null)
    {
        defaultXml.SetAttributeValue("ContentType", "application/xml");
    }

    var hasPresentationOverride = typesElement.Elements(contentTypesNs + "Override")
        .Any(e =>
            string.Equals((string?)e.Attribute("PartName"), "/ppt/presentation.xml", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(
                (string?)e.Attribute("ContentType"),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                StringComparison.OrdinalIgnoreCase));

    if (!hasPresentationOverride)
    {
        typesElement.Add(new XElement(
            contentTypesNs + "Override",
            new XAttribute("PartName", "/ppt/presentation.xml"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")));
    }

    entry.Delete();
    var replacementEntry = zip.CreateEntry("[Content_Types].xml");
    using var outputStream = replacementEntry.Open();
    document.Save(outputStream);
}
