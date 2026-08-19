@echo off
setlocal

echo ===================================================
echo   IIDelta Medical Writing Tools Installer
echo ===================================================
echo.

set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "SOURCE_FILE=%~dp0dist\IIDelta_MW_Tools.dotm"

tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [WARNING] Microsoft Word is currently running.
    echo Please close Word before installing the update.
    echo.
    pause
    exit /b
)

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
echo [SUCCESS] Tools installed successfully! Open Word to see the new Ribbon tab.
echo.
pause