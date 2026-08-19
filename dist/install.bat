@echo off
color 0A
title Medical Writing Tools - Installer

echo ===================================================
echo   Medical Writing Tools - Installation Script
echo ===================================================
echo.

:: Check if the .dotm file is in the same folder as the script
if not exist "%~dp0MedicalWritingTools.dotm" (
    echo [ERROR] Could not find "MedicalWritingTools.dotm"
    echo Please make sure this script and the .dotm file are in the same folder!
    echo.
    pause
    exit /b 1
)

:: Create the STARTUP folder if it doesn't exist just in case
if not exist "%APPDATA%\Microsoft\Word\STARTUP" (
    mkdir "%APPDATA%\Microsoft\Word\STARTUP"
)

:: Copy the file to the Word STARTUP folder
echo Copying tools to your Microsoft Word STARTUP folder...
copy /Y "%~dp0MedicalWritingTools.dotm" "%APPDATA%\Microsoft\Word\STARTUP\" >nul

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Installation failed. Microsoft Word might be locking the file.
    echo Please completely close Microsoft Word and Outlook, then try again.
    echo.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Installation Complete!
echo.
echo You can now open Microsoft Word. You will find your new macros
echo under the "Add-Ins" tab on the ribbon.
echo.
pause
