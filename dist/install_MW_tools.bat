@echo off
setlocal

echo ===================================================
echo   Medical Writing Tools Installer
echo ===================================================
echo.

set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "SOURCE_FILE=MedicalWritingTools.dotm"


if not exist "%SOURCE_FILE%" (
    echo [ERROR] Could not find the .dotm file in the dist/ folder.
    echo.
    pause
    exit /b
)

if not exist "%STARTUP_DIR%" mkdir "%STARTUP_DIR%"

echo Installing to: %STARTUP_DIR%
copy /Y "%SOURCE_FILE%" "%STARTUP_DIR%\" >NUL

echo.
echo [SUCCESS] Tools installed successfully! Remember to restart Word if it is currently open. Re-open Word to see the new 'Add-Ins' ribbon tab.
echo.
pause