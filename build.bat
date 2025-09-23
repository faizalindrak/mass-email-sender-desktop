# Build script for PyInstaller
@echo off
echo Building Email Automation Desktop...

REM Clean previous builds
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

REM Create executable
pyinstaller email_automation.spec

REM Check if build was successful
if exist "dist\EmailAutomation.exe" (
    echo.
    echo Build completed successfully!
    echo Executable: dist\EmailAutomation.exe
    echo Directory distribution: dist\EmailAutomation_dist\
    echo.
    echo You can also run the directory version which may start faster:
    echo dist\EmailAutomation_dist\EmailAutomation.exe
) else (
    echo.
    echo Build failed! Check the output above for errors.
)

pause