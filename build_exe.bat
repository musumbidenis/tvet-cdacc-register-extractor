@echo off
echo ============================================================
echo  TVET CDACC Register Extractor -- EXE Build Script
echo ============================================================
echo.

REM Make sure PyInstaller is installed
python -m pip install --quiet pyinstaller

REM Clean previous build artefacts
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

echo Building EXE (this takes 1-3 minutes)...
echo.
pyinstaller register_extractor.spec

echo.
if exist "dist\Register Extractor.exe" (
    echo ============================================================
    echo  SUCCESS!
    echo  EXE location: dist\Register Extractor.exe
    echo  Share that single file -- no Python needed on target PC.
    echo ============================================================
) else (
    echo ============================================================
    echo  BUILD FAILED -- check the output above for errors.
    echo ============================================================
)
pause
