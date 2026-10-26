# Photo Randomise and PowerPoint Maker

A .NET console app that:
- scans `PhotosInput` (including subfolders),
- copies supported photos into `PhotosOutput` with random GUID filenames,
- skips duplicate photos with identical file contents,
- ignores video files (including `.mp4`),
- generates a `Photos.pptx` slide deck from the copied images.

## Requirements

- .NET SDK 10.0+
- macOS, Windows, or Linux

## Supported input photo formats

`.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.webp`, `.heic`, `.heif`

## Ignored video formats

`.mp4`, `.mov`, `.avi`, `.mkv`, `.wmv`, `.flv`, `.webm`, `.m4v`, `.mpeg`, `.mpg`, `.3gp`

## PowerPoint slide image formats

PowerPoint slides are created from image types supported by the app's OpenXML pipeline:
- `.jpg` / `.jpeg`
- `.png`
- `.gif`
- `.bmp`
- `.tif` / `.tiff`

On macOS, HEIC/HEIF images can be converted to JPEG for the presentation when needed.

## Run

From repository root:

```bash
dotnet run --project PhotoRandomizerApp/PhotoRandomizerApp.csproj
```

Optional custom folders:

```bash
dotnet run --project PhotoRandomizerApp/PhotoRandomizerApp.csproj -- "/path/to/input" "/path/to/output"
```

If no arguments are provided, defaults are:
- Input: `./PhotosInput`
- Output: `./PhotosOutput`

## Output

After execution, the app reports:
- how many photos were copied,
- how many duplicate photos were skipped (if any),
- how many video files were ignored (if any),
- whether `Photos.pptx` was created and number of slides.
