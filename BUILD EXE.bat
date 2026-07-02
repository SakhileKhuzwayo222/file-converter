@echo off
setlocal
cd /d "%~dp0"

if not exist ".build-venv\Scripts\python.exe" (
    python -m venv .build-venv
)

".build-venv\Scripts\python.exe" -m pip install --upgrade pip pyinstaller
if errorlevel 1 (
    echo.
    echo Could not install PyInstaller. Check your internet connection, then run this file again.
    pause
    exit /b 1
)

".build-venv\Scripts\python.exe" tools\build_exe.py
if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done. Your executable is in:
echo %CD%\dist\File Converter.exe
pause
