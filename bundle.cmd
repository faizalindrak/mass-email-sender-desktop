@echo off
rem Bundle Windows executable for Email Automation Desktop using PyInstaller
setlocal enabledelayedexpansion

set APP_NAME=EmailAutomation
set SPEC_FILE=email_automation.spec
set VENV_DIR=venv
set PYTHON=python

rem Prefer virtual environment Python if available
if exist "%VENV_DIR%\Scripts\python.exe" (
    set PYTHON="%VENV_DIR%\Scripts\python.exe"
    call "%VENV_DIR%\Scripts\activate.bat"
) else (
    echo Virtual environment not found, using system Python.
)

rem Optional flags:
rem   --deps    Install requirements before building
rem   --release Zip the directory distribution to release\%APP_NAME%_dist.zip
set INSTALL_DEPS=0
set MAKE_RELEASE=0
for %%A in (%*) do (
    if /I "%%~A"=="--deps" set INSTALL_DEPS=1
    if /I "%%~A"=="--release" set MAKE_RELEASE=1
)

echo.
echo Detecting Python version and architecture...
%PYTHON% -c "import platform; print('Python %s (%s)') % (platform.python_version(), platform.architecture()[0])"
if errorlevel 1 (
    echo Failed to run Python. Ensure Python is installed and added to PATH.
    goto :end_fail
)

if %INSTALL_DEPS%==1 (
    echo.
    echo Installing project dependencies...
    %PYTHON% -m pip install --upgrade pip
    if errorlevel 1 goto :end_fail
    %PYTHON% -m pip install -r requirements.txt
    if errorlevel 1 goto :end_fail
)

echo.
echo Ensuring PyInstaller is installed...
%PYTHON% -m pip show pyinstaller >nul 2&gt;&1
if errorlevel 1 (
    echo Installing PyInstaller...
    %PYTHON% -m pip install pyinstaller
    if errorlevel 1 goto :end_fail
)

echo.
echo Cleaning previous build artifacts...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
mkdir dist >nul 2&gt;&1

echo.
echo Building with spec: %SPEC_FILE%
pyinstaller --noconfirm --clean "%SPEC_FILE%"
if errorlevel 1 goto :end_fail

echo.
echo Verifying build outputs...
set EXE_OK=
set DIR_OK=
if exist "dist\%APP_NAME%.exe" set EXE_OK=1
if exist "dist\%APP_NAME%_dist\%APP_NAME%.exe" set DIR_OK=1

if defined EXE_OK (
    echo Single-file executable: dist\%APP_NAME%.exe
)
if defined DIR_OK (
    echo Directory distribution: dist\%APP_NAME%_dist\%APP_NAME%.exe
)

if not defined EXE_OK if not defined DIR_OK goto :end_fail

if %MAKE_RELEASE%==1 (
    echo.
    echo Creating release zip from directory distribution...
    if not exist "release" mkdir "release"
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Compress-Archive -Path 'dist\\%APP_NAME%_dist\\*' -DestinationPath 'release\\%APP_NAME%_dist.zip' -Force"
    if errorlevel 1 (
        echo Failed to create release zip. Skipping compression.
    ) else (
        echo Release archive created: release\%APP_NAME%_dist.zip
    )
)

echo.
echo Build completed successfully.
echo You can run:
echo   dist\%APP_NAME%.exe          (single-file)
echo   dist\%APP_NAME%_dist\%APP_NAME%.exe  (faster startup)
goto :end_ok

:end_fail
echo.
echo Build failed. Review the errors above.
goto :end

:end_ok
echo.
echo Done.

:end
pause