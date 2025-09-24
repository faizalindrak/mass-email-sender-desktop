#Requires -Version 5.1

param(
    [switch]$Clean,
    [switch]$InstallDeps,
    [switch]$MakeRelease,
    [switch]$Portable,
    [switch]$NoUPX,
    [switch]$Debug,
    [switch]$Help
)

# Configuration
$AppName = "EmailAutomation"
$AppVersion = "1.0.0"
$SpecFile = "email_automation_optimized.spec"
$VenvDir = "venv"
$Python = "python"
$RequirementsFile = "requirements_optimized.txt"

# Colors for output
$ColorInfo = "Cyan"
$ColorSuccess = "Green" 
$ColorWarning = "Yellow"
$ColorError = "Red"

# Display help
if ($Help) {
    Write-Host "Email Automation Desktop - Optimized Build Script" -ForegroundColor $ColorInfo
    Write-Host "====================================================" -ForegroundColor $ColorInfo
    Write-Host ""
    Write-Host "Usage: .\build.ps1 [options]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -Clean          Clean all build artifacts before building"
    Write-Host "  -InstallDeps    Install optimized dependencies before building"
    Write-Host "  -MakeRelease    Create release packages (zip and installer)"
    Write-Host "  -Portable       Create portable version (directory distribution)"
    Write-Host "  -NoUPX          Disable UPX compression for faster debugging"
    Write-Host "  -Debug          Build with debug console enabled"
    Write-Host "  -Help           Show this help message"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\build.ps1                           # Quick build"
    Write-Host "  .\build.ps1 -Clean -InstallDeps       # Clean build with deps"
    Write-Host "  .\build.ps1 -MakeRelease -Portable    # Full release build"
    Write-Host "  .\build.ps1 -Debug -NoUPX             # Debug build"
    Write-Host ""
    exit 0
}

Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host "Email Automation Desktop - Optimized Build" -ForegroundColor $ColorSuccess
Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host ""

# Function to check command availability
function Test-CommandExists {
    param($Command)
    try {
        Get-Command $Command -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Function to show progress
function Show-Progress {
    param($Activity, $Status)
    Write-Host "[$Activity] $Status" -ForegroundColor $ColorInfo
}

# Check Python environment
Show-Progress "SETUP" "Checking Python environment..."

# Detect virtual environment
if (Test-Path "$VenvDir\Scripts\python.exe") {
    $Python = "$VenvDir\Scripts\python.exe"
    Write-Host "Using virtual environment: $VenvDir" -ForegroundColor $ColorWarning
    
    # Activate virtual environment
    try {
        & "$VenvDir\Scripts\Activate.ps1"
        Write-Host "Virtual environment activated" -ForegroundColor $ColorSuccess
    } catch {
        Write-Host "Warning: Failed to activate virtual environment" -ForegroundColor $ColorWarning
    }
} else {
    Write-Host "Using system Python" -ForegroundColor $ColorWarning
    if (-not (Test-CommandExists "python")) {
        Write-Host "ERROR: Python not found in PATH" -ForegroundColor $ColorError
        Write-Host "Please install Python 3.9+ or create a virtual environment" -ForegroundColor $ColorError
        exit 1
    }
}

# Check Python version
Show-Progress "SETUP" "Validating Python version..."
try {
    $PythonInfo = & $Python -c "import sys; print(f'Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} ({sys.platform})')"
    Write-Host $PythonInfo -ForegroundColor $ColorSuccess
    
    $VersionCheck = & $Python -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Python 3.9+ is required" -ForegroundColor $ColorError
        exit 1
    }
} catch {
    Write-Host "ERROR: Failed to check Python version" -ForegroundColor $ColorError
    exit 1
}

# Clean build artifacts
if ($Clean) {
    Show-Progress "CLEAN" "Removing build artifacts..."
    
    $DirsToClean = @("build", "dist", "release", "__pycache__")
    foreach ($dir in $DirsToClean) {
        if (Test-Path $dir) {
            Write-Host "Removing $dir..." -ForegroundColor $ColorWarning
            Remove-Item -Path $dir -Recurse -Force
        }
    }
    
    # Clean Python cache files
    Get-ChildItem -Recurse -Name "*.pyc" | Remove-Item -Force
    Get-ChildItem -Recurse -Name "__pycache__" | Remove-Item -Recurse -Force
    
    Write-Host "Build artifacts cleaned" -ForegroundColor $ColorSuccess
}

# Install dependencies
if ($InstallDeps) {
    Show-Progress "DEPS" "Installing optimized dependencies..."
    
    if (-not (Test-Path $RequirementsFile)) {
        Write-Host "ERROR: $RequirementsFile not found" -ForegroundColor $ColorError
        exit 1
    }
    
    # Check for uv (faster package installer)
    $UseUv = Test-CommandExists "uv"
    if ($UseUv) {
        Write-Host "Using uv for faster installation..." -ForegroundColor $ColorWarning
        & uv pip install -p $Python -r $RequirementsFile --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Failed to install dependencies with uv" -ForegroundColor $ColorError
            exit 1
        }
    } else {
        Write-Host "Using pip..." -ForegroundColor $ColorWarning
        & $Python -m pip install -r $RequirementsFile --quiet --upgrade
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Failed to install dependencies with pip" -ForegroundColor $ColorError
            exit 1
        }
    }
    
    Write-Host "Dependencies installed successfully" -ForegroundColor $ColorSuccess
}

# Ensure PyInstaller is available
Show-Progress "DEPS" "Verifying PyInstaller..."
$PyInstallerCheck = & $Python -c "import PyInstaller; print('PyInstaller', PyInstaller.__version__)" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing PyInstaller..." -ForegroundColor $ColorWarning
    & $Python -m pip install pyinstaller --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to install PyInstaller" -ForegroundColor $ColorError
        exit 1
    }
} else {
    Write-Host $PyInstallerCheck -ForegroundColor $ColorSuccess
}

# Create directories
New-Item -ItemType Directory -Path "dist" -Force | Out-Null

# Update spec file for current build options
Show-Progress "BUILD" "Preparing build configuration..."

$SpecContent = Get-Content $SpecFile -Raw
if ($Debug) {
    $SpecContent = $SpecContent -replace "console=False", "console=True"
    Write-Host "Debug console enabled" -ForegroundColor $ColorWarning
} else {
    $SpecContent = $SpecContent -replace "console=True", "console=False"
}

if ($NoUPX) {
    $SpecContent = $SpecContent -replace "upx=True", "upx=False"
    Write-Host "UPX compression disabled" -ForegroundColor $ColorWarning
} else {
    $SpecContent = $SpecContent -replace "upx=False", "upx=True"
}

Set-Content -Path $SpecFile -Value $SpecContent

# Build executable
Show-Progress "BUILD" "Building executable with PyInstaller..."
Write-Host "This may take several minutes..." -ForegroundColor $ColorWarning

$BuildArgs = @("--noconfirm", "--clean")
if ($Debug) {
    $BuildArgs += "--debug=all"
}

& $Python -m PyInstaller @BuildArgs $SpecFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: PyInstaller build failed" -ForegroundColor $ColorError
    Write-Host "Check the error messages above for details" -ForegroundColor $ColorError
    exit 1
}

# Verify build outputs
Show-Progress "VERIFY" "Checking build outputs..."

$SingleFileExe = "dist\$AppName.exe"
$DirDistExe = "dist\$AppName`_dist\$AppName.exe"

$SingleFileOk = Test-Path $SingleFileExe
$DirDistOk = Test-Path $DirDistExe

if ($SingleFileOk) {
    $Size = [math]::Round((Get-Item $SingleFileExe).Length / 1MB, 1)
    Write-Host "✓ Single-file executable: $SingleFileExe ($Size MB)" -ForegroundColor $ColorSuccess
}

if ($DirDistOk) {
    $DirSize = [math]::Round((Get-ChildItem "dist\$AppName`_dist" -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
    Write-Host "✓ Directory distribution: $DirDistExe ($DirSize MB)" -ForegroundColor $ColorSuccess
}

if (-not $SingleFileOk -and -not $DirDistOk) {
    Write-Host "ERROR: No executable files found in dist/" -ForegroundColor $ColorError
    exit 1
}

# Test executable
Show-Progress "TEST" "Testing executable..."
if ($SingleFileOk) {
    try {
        $TestResult = Start-Process -FilePath $SingleFileExe -ArgumentList "--help" -Wait -PassThru -WindowStyle Hidden
        if ($TestResult.ExitCode -eq 0) {
            Write-Host "✓ Executable test passed" -ForegroundColor $ColorSuccess
        } else {
            Write-Host "⚠ Executable test returned non-zero exit code" -ForegroundColor $ColorWarning
        }
    } catch {
        Write-Host "⚠ Could not test executable (may require GUI)" -ForegroundColor $ColorWarning
    }
}

# Create release packages
if ($MakeRelease -or $Portable) {
    Show-Progress "RELEASE" "Creating release packages..."
    
    if (-not (Test-Path "release")) {
        New-Item -ItemType Directory -Path "release" -Force | Out-Null
    }
    
    $ReleaseDate = Get-Date -Format "yyyy-MM-dd"
    $ReleaseName = "$AppName-v$AppVersion-$ReleaseDate"
    
    # Create portable version
    if ($Portable -and $DirDistOk) {
        $PortableName = "$ReleaseName-portable"
        Write-Host "Creating portable version: $PortableName.zip" -ForegroundColor $ColorWarning
        
        try {
            Compress-Archive -Path "dist\$AppName`_dist\*" -DestinationPath "release\$PortableName.zip" -Force
            $ZipSize = [math]::Round((Get-Item "release\$PortableName.zip").Length / 1MB, 1)
            Write-Host "✓ Portable version created: $PortableName.zip ($ZipSize MB)" -ForegroundColor $ColorSuccess
        } catch {
            Write-Host "⚠ Failed to create portable version" -ForegroundColor $ColorWarning
        }
    }
    
    # Create single-file release
    if ($MakeRelease -and $SingleFileOk) {
        $SingleFileName = "$ReleaseName-standalone"
        Write-Host "Creating standalone version: $SingleFileName.zip" -ForegroundColor $ColorWarning
        
        try {
            # Create temp directory with additional files
            $TempDir = "temp_release"
            New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
            
            # Copy executable
            Copy-Item $SingleFileExe -Destination "$TempDir\$AppName.exe"
            
            # Copy additional files
            $AdditionalFiles = @("README.md", "LICENSE", "config", "templates")
            foreach ($file in $AdditionalFiles) {
                if (Test-Path $file) {
                    Copy-Item $file -Destination $TempDir -Recurse -Force
                }
            }
            
            # Create release zip
            Compress-Archive -Path "$TempDir\*" -DestinationPath "release\$SingleFileName.zip" -Force
            
            # Cleanup
            Remove-Item -Path $TempDir -Recurse -Force
            
            $ZipSize = [math]::Round((Get-Item "release\$SingleFileName.zip").Length / 1MB, 1)
            Write-Host "✓ Standalone release created: $SingleFileName.zip ($ZipSize MB)" -ForegroundColor $ColorSuccess
        } catch {
            Write-Host "⚠ Failed to create standalone release" -ForegroundColor $ColorWarning
        }
    }
}

# Build summary
Write-Host ""
Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host "Build completed successfully!" -ForegroundColor $ColorSuccess
Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host ""

Write-Host "Available executables:" -ForegroundColor $ColorInfo
if ($SingleFileOk) {
    Write-Host "  • $SingleFileExe (single-file, slower startup)" -ForegroundColor "White"
}
if ($DirDistOk) {
    Write-Host "  • $DirDistExe (directory, faster startup)" -ForegroundColor "White"
}

if ($MakeRelease -or $Portable) {
    Write-Host ""
    Write-Host "Release packages created in 'release/' directory" -ForegroundColor $ColorInfo
    Get-ChildItem "release\" -Name "*.zip" | ForEach-Object {
        Write-Host "  • $_" -ForegroundColor "White"
    }
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor $ColorInfo
Write-Host "  1. Test the executable on target systems" -ForegroundColor "White"
Write-Host "  2. Verify all features work correctly" -ForegroundColor "White"
Write-Host "  3. Distribute the appropriate package" -ForegroundColor "White"
Write-Host ""