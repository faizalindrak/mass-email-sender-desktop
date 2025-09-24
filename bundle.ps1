#Requires -Version 5.1

param(
    [switch]$InstallDeps,
    [switch]$MakeRelease,
    [switch]$Help
)

# Configuration
$AppName = "EmailAutomation"
$SpecFile = "email_automation.spec"
$VenvDir = "venv"
$Python = "python"

# Display help if requested
if ($Help) {
    Write-Host "Email Automation Desktop - PowerShell Build Script" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage: .\bundle.ps1 [options]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -InstallDeps    Install project dependencies before building"
    Write-Host "  -MakeRelease    Create release zip from directory distribution"
    Write-Host "  -Help           Show this help message"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\bundle.ps1                    # Build without installing deps"
    Write-Host "  .\bundle.ps1 -InstallDeps       # Install deps and build"
    Write-Host "  .\bundle.ps1 -MakeRelease       # Build and create release zip"
    Write-Host ""
    exit 0
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "Email Automation Desktop - Build Script" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if uv is available
$UseUv = $false
try {
    $null = Get-Command uv -ErrorAction Stop
    $UseUv = $true
    Write-Host "Found uv package manager" -ForegroundColor Yellow
} catch {
    Write-Host "uv not found, using pip" -ForegroundColor Yellow
}

# Check for virtual environment
if (Test-Path "$VenvDir\Scripts\python.exe") {
    $Python = "$VenvDir\Scripts\python.exe"
    Write-Host "Using virtual environment: $VenvDir" -ForegroundColor Yellow
    try {
        & "$VenvDir\Scripts\Activate.ps1"
    } catch {
        Write-Warning "Failed to activate virtual environment"
    }
} else {
    Write-Host "Virtual environment not found, using system Python" -ForegroundColor Yellow
    # Check if Python is available
    try {
        $null = Get-Command python -ErrorAction Stop
    } catch {
        Write-Error "Python is not installed or not in PATH"
        Write-Error "Please install Python 3.9+ and add it to your PATH"
        exit 1
    }
}

# Check Python version
Write-Host "Detecting Python version and architecture..." -ForegroundColor Cyan
try {
    $PythonVersion = & $Python -c "import platform; print(f'Python {platform.python_version()} ({platform.architecture()[0]})')"
    Write-Host $PythonVersion -ForegroundColor Green

    # Check Python version compatibility
    $VersionCheck = & $Python -c "import sys; v = sys.version_info; print(f'Python {v.major}.{v.minor}.{v.micro}'); sys.exit(0 if v >= (3, 9) else 1)"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Python 3.9+ is required. Please upgrade Python."
        exit 1
    }
} catch {
    Write-Error "Failed to run Python. Ensure Python is installed and added to PATH."
    exit 1
}

# Install dependencies if requested
if ($InstallDeps) {
    Write-Host "Installing project dependencies..." -ForegroundColor Cyan

    if ($UseUv) {
        Write-Host "Using uv to install dependencies..." -ForegroundColor Yellow
        & uv pip install -p $Python -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install dependencies with uv"
            exit 1
        }
    } else {
        Write-Host "Ensuring pip is available..." -ForegroundColor Yellow
        & $Python -m ensurepip --upgrade
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to setup pip"
            exit 1
        }

        Write-Host "Upgrading pip..." -ForegroundColor Yellow
        & $Python -m pip install --upgrade pip --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to upgrade pip"
            exit 1
        }

        Write-Host "Installing dependencies..." -ForegroundColor Yellow
        & $Python -m pip install -r requirements.txt --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install dependencies"
            exit 1
        }
    }

    Write-Host "Dependencies installed successfully." -ForegroundColor Green
}

# Ensure PyInstaller is installed
Write-Host "Ensuring PyInstaller is installed..." -ForegroundColor Cyan
try {
    $PyInstallerCheck = & $Python -c "import PyInstaller" 2>$null
    Write-Host "PyInstaller already installed." -ForegroundColor Green
} catch {
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    if ($UseUv) {
        & uv pip install -p $Python pyinstaller --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install PyInstaller with uv"
            exit 1
        }
    } else {
        & $Python -m pip install pyinstaller --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install PyInstaller"
            exit 1
        }
    }
    Write-Host "PyInstaller installed successfully." -ForegroundColor Green
}

# Clean previous build artifacts
Write-Host "Cleaning previous build artifacts..." -ForegroundColor Cyan
if (Test-Path "build") {
    Write-Host "Removing build directory..." -ForegroundColor Yellow
    Remove-Item -Path "build" -Recurse -Force
}
if (Test-Path "dist") {
    Write-Host "Removing dist directory..." -ForegroundColor Yellow
    Remove-Item -Path "dist" -Recurse -Force
}
if (Test-Path "release") {
    Write-Host "Removing release directory..." -ForegroundColor Yellow
    Remove-Item -Path "release" -Recurse -Force
}
New-Item -ItemType Directory -Path "dist" -Force | Out-Null

# Build with PyInstaller
Write-Host "Building executable with spec: $SpecFile" -ForegroundColor Cyan
Write-Host "This may take several minutes..." -ForegroundColor Yellow
& $Python -m PyInstaller --noconfirm --clean $SpecFile
if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller build failed"
    Write-Error "Check the error messages above for details"
    exit 1
}

# Verify build outputs
Write-Host "Verifying build outputs..." -ForegroundColor Cyan
$ExeOk = Test-Path "dist\$AppName.exe"
$DirOk = Test-Path "dist\$AppName`_dist\$AppName.exe"

if ($ExeOk) {
    Write-Host "Single-file executable: dist\$AppName.exe" -ForegroundColor Green
}
if ($DirOk) {
    Write-Host "Directory distribution: dist\$AppName`_dist\$AppName.exe" -ForegroundColor Green
}

if (-not $ExeOk -and -not $DirOk) {
    Write-Error "Build verification failed - no executable found"
    exit 1
}

# Create release zip if requested
if ($MakeRelease) {
    Write-Host "Creating release zip from directory distribution..." -ForegroundColor Cyan
    if (-not (Test-Path "release")) {
        New-Item -ItemType Directory -Path "release" -Force | Out-Null
    }

    try {
        Compress-Archive -Path "dist\$AppName`_dist\*" -DestinationPath "release\$AppName`_dist.zip" -Force
        Write-Host "Release archive created: release\$AppName`_dist.zip" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create release zip. Skipping compression."
    }
}

Write-Host ""
Write-Host "Build completed successfully!" -ForegroundColor Green
Write-Host "You can run:" -ForegroundColor Cyan
if ($ExeOk) {
    Write-Host "  dist\$AppName.exe          (single-file)" -ForegroundColor White
}
if ($DirOk) {
    Write-Host "  dist\$AppName`_dist\$AppName.exe  (faster startup)" -ForegroundColor White
}