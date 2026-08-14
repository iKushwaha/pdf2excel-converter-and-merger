#!/usr/bin/env bash
# ============================================================
#  Build a standalone executable for "PDF -> Excel Converter &
#  Merger" using PyInstaller.
#
#  Works on macOS, Linux and Windows (Git Bash). On Windows the
#  .exe is produced via scripts/build_windows.bat too.
#
#  Usage:
#    scripts/build.sh
#
#  Output:
#    macOS: dist/PDF2Excel.app (drag to Applications)
#    Linux: dist/PDF2Excel     (executable)
#    Windows: dist/PDF2Excel.exe
# ============================================================
set -euo pipefail
cd "$(dirname "$0")/.."

echo "[1/3] Installing dependencies (including PyInstaller)..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller

echo "[2/3] Building..."
python3 -m PyInstaller --noconfirm --clean --onefile --windowed \
  --name "PDF2Excel" \
  main.py

echo "[3/3] Build complete."
echo
echo "  dist/PDF2Excel$(if [[ "$OSTYPE" == "msys"* || "$OSTYPE" == "cygwin"* || "$OSTYPE" == "mingw"* ]]; then echo .exe; elif [[ "$(uname)" == "Darwin" ]]; then echo .app; fi)"
echo
echo "To publish the Windows build, copy dist/PDF2Excel.exe to"
echo "web/downloads/PDF2Excel-win64.exe and deploy the web/ folder."
