@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on your PATH.
    echo.
    echo Please install Python from https://python.org and make sure the box
    echo "Add python.exe to PATH" is checked during installation.
    echo.
    echo If you already installed Python, try closing this window and opening
    echo a new one - PATH changes only take effect in new windows.
    echo.
    pause
    exit /b 1
)

echo Found Python. Running shipment processing script...
echo.

python OutstandingShipments.py
set SCRIPT_EXIT=%ERRORLEVEL%

echo.
if %SCRIPT_EXIT% NEQ 0 (
    echo Script stopped - see the message above for what needs fixing.
) else (
    echo Script finished successfully.
)
echo.
pause
