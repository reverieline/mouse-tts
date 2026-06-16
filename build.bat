@echo off
setlocal
cd /d "%~dp0"

echo Installing / updating dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: pip install failed.
    exit /b 1
)

echo.
echo Building mouse_tts.exe ...
python -m PyInstaller mouse_tts.spec --noconfirm --distpath dist --workpath build
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete.
echo  Executable:  %~dp0dist\mouse_tts.exe
echo.
echo  Double-click dist\mouse_tts.exe to run.
echo  A settings window opens on first launch.
echo  Closing the window minimizes to the system tray.
echo ============================================================
