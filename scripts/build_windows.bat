@echo off
REM ============================================================
REM  Build a Windows executable for "PDF -> Excel Converter &
REM  Merger" using PyInstaller.
REM
REM  Requirements:
REM    - Windows 10/11 (or a Windows VM)
REM    - Python 3.10+ installed and on PATH
REM    - Run from the project root:  scripts\build_windows.bat
REM
REM  Output:
REM    dist\PDF2Excel.exe           (single file, no Python needed)
REM
REM  To publish:
REM    copy dist\PDF2Excel.exe web\downloads\PDF2Excel-win64.exe
REM    (then deploy the web\ folder to any static host)
REM ============================================================
setlocal
cd /d "%~dp0\.."

echo [1/3] Installing dependencies (including PyInstaller)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

echo [2/3] Building the executable...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "PDF2Excel" ^
  --icon web\assets\app.ico ^
  main.py
if errorlevel 1 goto :error

echo [3/3] Build complete.
echo.
echo   dist\PDF2Excel.exe
echo.
echo Optional: wrap dist\PDF2Excel.exe with NSIS or Inno Setup
echo to produce a Setup.exe with Start Menu shortcuts and an icon.
echo Then copy the result to:  web\downloads\PDF2Excel-win64.exe
exit /b 0

:error
echo.
echo BUILD FAILED. See the messages above.
exit /b 1
