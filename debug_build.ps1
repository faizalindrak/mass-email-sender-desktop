#Requires -Version 5.1

param(
    [switch]$Clean,
    [switch]$InstallDeps,
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

if ($Help) {
    Write-Host "Email Automation Desktop - Debug Build Script" -ForegroundColor $ColorInfo
    Write-Host "===============================================" -ForegroundColor $ColorInfo
    Write-Host ""
    Write-Host "This script builds the application with debug console enabled"
    Write-Host "to help troubleshoot runtime issues."
    Write-Host ""
    Write-Host "Usage: .\debug_build.ps1 [options]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -Clean          Clean all build artifacts before building"
    Write-Host "  -InstallDeps    Install dependencies before building"  
    Write-Host "  -Help           Show this help message"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\debug_build.ps1                    # Debug build"
    Write-Host "  .\debug_build.ps1 -Clean -InstallDeps # Clean debug build with deps"
    Write-Host ""
    exit 0
}

Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host "Email Automation Desktop - Debug Build" -ForegroundColor $ColorSuccess
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

# Check Python environment
Write-Host "[SETUP] Checking Python environment..." -ForegroundColor $ColorInfo

# Detect virtual environment
if (Test-Path "$VenvDir\Scripts\python.exe") {
    $Python = "$VenvDir\Scripts\python.exe"
    Write-Host "Using virtual environment: $VenvDir" -ForegroundColor $ColorWarning
    
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
        exit 1
    }
}

# Check Python version
Write-Host "[SETUP] Validating Python version..." -ForegroundColor $ColorInfo
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

# Clean build artifacts if requested
if ($Clean) {
    Write-Host "[CLEAN] Removing build artifacts..." -ForegroundColor $ColorInfo
    
    $DirsToClean = @("build", "dist", "__pycache__")
    foreach ($dir in $DirsToClean) {
        if (Test-Path $dir) {
            Write-Host "Removing $dir..." -ForegroundColor $ColorWarning
            Remove-Item -Path $dir -Recurse -Force
        }
    }
    
    # Clean Python cache files
    Get-ChildItem -Recurse -Name "*.pyc" -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -Recurse -Name "__pycache__" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    
    Write-Host "Build artifacts cleaned" -ForegroundColor $ColorSuccess
}

# Install dependencies if requested
if ($InstallDeps) {
    Write-Host "[DEPS] Installing dependencies..." -ForegroundColor $ColorInfo
    
    if (-not (Test-Path $RequirementsFile)) {
        Write-Host "ERROR: $RequirementsFile not found" -ForegroundColor $ColorError
        exit 1
    }
    
    # Check for uv (faster package installer)
    $UseUv = Test-CommandExists "uv"
    if ($UseUv) {
        Write-Host "Using uv for installation..." -ForegroundColor $ColorWarning
        & uv pip install -p $Python -r $RequirementsFile
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Failed to install dependencies with uv" -ForegroundColor $ColorError
            exit 1
        }
    } else {
        Write-Host "Using pip..." -ForegroundColor $ColorWarning
        & $Python -m pip install -r $RequirementsFile --upgrade
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Failed to install dependencies with pip" -ForegroundColor $ColorError
            exit 1
        }
    }
    
    Write-Host "Dependencies installed" -ForegroundColor $ColorSuccess
}

# Ensure PyInstaller
Write-Host "[DEPS] Checking PyInstaller..." -ForegroundColor $ColorInfo
$PyInstallerCheck = & $Python -c "import PyInstaller; print('PyInstaller', PyInstaller.__version__)" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing PyInstaller..." -ForegroundColor $ColorWarning
    & $Python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to install PyInstaller" -ForegroundColor $ColorError
        exit 1
    }
} else {
    Write-Host $PyInstallerCheck -ForegroundColor $ColorSuccess
}

# Test imports before building
Write-Host "[TEST] Testing critical imports..." -ForegroundColor $ColorInfo
$TestImports = @(
    "ast",
    "PySide6.QtCore",
    "PySide6.QtWidgets", 
    "qfluentwidgets",
    "watchdog",
    "jinja2",
    "win32com.client"
)

foreach ($module in $TestImports) {
    try {
        $TestResult = & $Python -c "import $module; print('✓ $module')" 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ $module" -ForegroundColor $ColorSuccess
        } else {
            Write-Host "✗ $module - FAILED" -ForegroundColor $ColorError
        }
    } catch {
        Write-Host "✗ $module - ERROR" -ForegroundColor $ColorError
    }
}

# Create directories
New-Item -ItemType Directory -Path "dist" -Force | Out-Null

# Update spec file for debug mode
Write-Host "[BUILD] Preparing debug build configuration..." -ForegroundColor $ColorInfo
$SpecContent = Get-Content $SpecFile -Raw
$DebugSpecContent = $SpecContent -replace "console=False", "console=True"
$DebugSpecContent = $DebugSpecContent -replace "debug=False", "debug=True"
$DebugSpecContent = $DebugSpecContent -replace "upx=True", "upx=False"

$DebugSpecFile = "debug_" + $SpecFile
Set-Content -Path $DebugSpecFile -Value $DebugSpecContent
Write-Host "Debug spec file created: $DebugSpecFile" -ForegroundColor $ColorSuccess

# Build with debug settings
Write-Host "[BUILD] Building with debug console enabled..." -ForegroundColor $ColorInfo
Write-Host "This will create an executable with console output for debugging" -ForegroundColor $ColorWarning
Write-Host ""

& $Python -m PyInstaller --noconfirm --clean $DebugSpecFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Debug build failed" -ForegroundColor $ColorError
    exit 1
}

# Check outputs
Write-Host "[VERIFY] Checking build outputs..." -ForegroundColor $ColorInfo
$DebugExe = "dist\$AppName.exe"
$DebugDist = "dist\$AppName`_dist\$AppName.exe"

if (Test-Path $DebugExe) {
    $Size = [math]::Round((Get-Item $DebugExe).Length / 1MB, 1)
    Write-Host "✓ Debug executable: $DebugExe ($Size MB)" -ForegroundColor $ColorSuccess
}

if (Test-Path $DebugDist) {
    $DirSize = [math]::Round((Get-ChildItem "dist\$AppName`_dist" -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
    Write-Host "✓ Debug directory: $DebugDist ($DirSize MB)" -ForegroundColor $ColorSuccess
}

# Quick test
Write-Host "[TEST] Testing debug executable..." -ForegroundColor $ColorInfo
if (Test-Path $DebugExe) {
    Write-Host "Starting debug test (will show console window)..." -ForegroundColor $ColorWarning
    Write-Host "Press Ctrl+C to stop if application hangs" -ForegroundColor $ColorWarning
    
    try {
        $TestProcess = Start-Process -FilePath $DebugExe -PassThru -WindowStyle Normal
        Start-Sleep -Seconds 3
        
        if (!$TestProcess.HasExited) {
            Write-Host "✓ Application started successfully (PID: $($TestProcess.Id))" -ForegroundColor $ColorSuccess
            Write-Host "Stopping test process..." -ForegroundColor $ColorWarning
            $TestProcess.CloseMainWindow()
            Start-Sleep -Seconds 2
            if (!$TestProcess.HasExited) {
                $TestProcess.Kill()
            }
        } else {
            Write-Host "✗ Application exited immediately (Exit Code: $($TestProcess.ExitCode))" -ForegroundColor $ColorError
        }
    } catch {
        Write-Host "✗ Failed to test executable: $($_.Exception.Message)" -ForegroundColor $ColorError
    }
}

# Cleanup debug spec
Remove-Item -Path $DebugSpecFile -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host "Debug build completed!" -ForegroundColor $ColorSuccess
Write-Host "=============================================" -ForegroundColor $ColorSuccess
Write-Host ""
Write-Host "Debug executable created with console output enabled." -ForegroundColor $ColorInfo
Write-Host "Run the executable from command line to see debug messages." -ForegroundColor $ColorInfo
Write-Host ""
Write-Host "If issues persist:" -ForegroundColor $ColorWarning
Write-Host "1. Run: dist\$AppName.exe" -ForegroundColor "White"
Write-Host "2. Check console output for specific error messages" -ForegroundColor "White" 
Write-Host "3. Verify all dependencies are properly installed" -ForegroundColor "White"
Write-Host "4. Try directory distribution: dist\$AppName`_dist\$AppName.exe" -ForegroundColor "White"
Write-Host ""