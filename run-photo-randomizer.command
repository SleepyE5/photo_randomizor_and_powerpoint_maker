#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "$0")" && pwd)"
APP_PATH="$SCRIPT_DIR/PhotoRandomizerApp/bin/Release/net10.0/osx-arm64/publish/PhotoRandomizerApp"

if [[ ! -x "$APP_PATH" ]]; then
  echo "Mac executable not found."
  echo "Build it first with:"
  echo "dotnet publish \"$SCRIPT_DIR/PhotoRandomizerApp/PhotoRandomizerApp.csproj\" -c Release -r osx-arm64 --self-contained true -p:PublishSingleFile=true"
  exit 1
fi

cd "$SCRIPT_DIR"

if [[ $# -eq 0 ]]; then
  "$APP_PATH" "$SCRIPT_DIR/PhotosInput" "$SCRIPT_DIR/PhotosOutput"
else
  "$APP_PATH" "$@"
fi
