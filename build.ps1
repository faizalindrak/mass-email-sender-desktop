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

# PyInstaller configuration details
$HiddenImports = @(
    # Python built-in modules (critical)
    'ast',
    'dis',
    'inspect',
    'copy',
    'pickle',
    'struct',
    'operator',
    'weakref',
    'gc',
    'io',
    'codecs',
    'encodings',
    'encodings.utf_8',
    'encodings.cp1252',
    'locale',
    'warnings',
    'linecache',
    'keyword',
    'token',
    'tokenize',

    # PySide6 core modules
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'PySide6.QtNetwork',
    'PySide6.QtSvg',
    'PySide6.QtPrintSupport',

    # Shiboken6
    'shiboken6',

    # Fluent widgets
    'qfluentwidgets',
    'qfluentwidgets.common',
    'qfluentwidgets.common.config',
    'qfluentwidgets.common.icon',
    'qfluentwidgets.common.style_sheet',
    'qfluentwidgets.components',
    'qfluentwidgets.components.widgets',
    'qfluentwidgets.components.layout',
    'qfluentwidgets.components.dialog_box',
    'qfluentwidgets.components.material',
    'qfluentwidgets.window',
    'qfluentwidgets._rc',

    # Windows COM for Outlook
    'win32com.client',
    'win32com.client.gencache',
    'win32com.client.CLSIDToClass',
    'win32com.client.util',
    'win32com.server',
    'win32com.server.util',
    'pythoncom',
    'pywintypes',
    'win32api',
    'win32con',
    'win32gui',
    'win32process',

    # File monitoring
    'watchdog',
    'watchdog.observers',
    'watchdog.observers.winapi',
    'watchdog.events',
    'watchdog.utils',

    # Template engine
    'jinja2',
    'jinja2.ext',
    'jinja2.loaders',
    'jinja2.runtime',
    'jinja2.compiler',
    'jinja2.environment',
    'markupsafe',

    # Email modules
    'smtplib',
    'email',
    'email.mime',
    'email.mime.multipart',
    'email.mime.text',
    'email.mime.base',
    'email.mime.application',
    'email.mime.message',
    'email.encoders',
    'email.utils',
    'email.header',
    'email.charset',

    # Database
    'sqlite3',

    # Configuration
    'configparser',

    # Logging
    'logging',
    'logging.handlers',
    'logging.config',

    # Standard library essentials
    'pathlib',
    'contextlib',
    'functools',
    'itertools',
    'collections',
    'collections.abc',
    'typing',
    'json',
    'os',
    'sys',
    're',
    'datetime',
    'time',
    'threading',
    'queue',
    'traceback',
    'base64',
    'uuid',
    'hashlib',
    'hmac',
    'urllib',
    'urllib.parse',
    'urllib.request',
    'urllib.error',
    'ssl',
    'socket',

    # Package management (for runtime)
    'pkg_resources',
    'importlib',
    'importlib.util',
    'importlib.metadata',

    # Theme detection
    'darkdetect'
)

$ExcludedModules = @(
    # Only exclude really unnecessary GUI frameworks
    'tkinter',
    'PyQt5',
    'PyQt6',
    'wx',
    'kivy',
    'toga',

    # Only exclude the heaviest scientific libraries
    'numpy',
    'pandas',
    'scipy',
    'matplotlib',
    'seaborn',
    'plotly',
    'bokeh',
    'altair',

    # Exclude development/testing tools only
    'pytest',
    'unittest',
    'doctest',
    'pdb',

    # Exclude Jupyter/IPython
    'IPython',
    'jupyter',
    'notebook',
    'jupyterlab',
    'ipykernel',
    'ipywidgets',

    # Exclude unused web frameworks
    'django',
    'flask',
    'fastapi',
    'tornado',
    'pyramid',
    'bottle',

    # Exclude unused databases
    'psycopg2',
    'pymongo',
    'redis',
    'mysql',
    'postgresql',

    # Exclude specific unused modules only
    'curses',
    'readline'
)

$DataDirectories = @(
    @{ Source = "templates"; Target = "templates" },
    @{ Source = "config"; Target = "config" },
    @{ Source = "database"; Target = "database" },
    @{ Source = "logs"; Target = "logs" }
)

$CollectPackages = @(
    "qfluentwidgets"
)

function Get-QFluentResourceMappings {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PythonExecutable
    )

    $script = @"
import json
import os

try:
    import qfluentwidgets
    base_path = os.path.dirname(qfluentwidgets.__file__)
    resource_dirs = []
    for dirname in ['qss', 'resources', '_rc']:
        path = os.path.join(base_path, dirname)
        if os.path.exists(path):
            resource_dirs.append([path, f'qfluentwidgets/{dirname}'])
    print(json.dumps({'resources': resource_dirs}))
except Exception as exc:
    print(json.dumps({'error': str(exc)}))
"@

    try {
        $raw = & $PythonExecutable -c $script 2>$null
    } catch {
        return @()
    }

    if (-not $raw) {
        return @()
    }

    try {
        $parsed = $raw | ConvertFrom-Json
    } catch {
        return @()
    }

    if ($parsed.error) {
        Write-Host "Warning: Failed to resolve qfluentwidgets resources: $($parsed.error)" -ForegroundColor $ColorWarning
        return @()
    }

    return $parsed.resources
}

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

# Create output directories
New-Item -ItemType Directory -Path "dist" -Force | Out-Null
New-Item -ItemType Directory -Path "build" -Force | Out-Null

Show-Progress "BUILD" "Preparing build configuration..."

try {
    $SrcPath = (Resolve-Path "src").Path
} catch {
    Write-Host "ERROR: Could not resolve project source directory" -ForegroundColor $ColorError
    exit 1
}

# Ensure required data directories exist and gather add-data arguments
$AddDataArgs = @()
foreach ($dirEntry in $DataDirectories) {
    $sourcePath = $null

    if ($dirEntry.Source -eq "logs" -and -not (Test-Path $dirEntry.Source)) {
        New-Item -ItemType Directory -Path $dirEntry.Source -Force | Out-Null
    }

    if (Test-Path $dirEntry.Source) {
        $sourcePath = (Resolve-Path $dirEntry.Source).Path
        $AddDataArgs += "--add-data=$sourcePath;$($dirEntry.Target)"
        Write-Host "Including data directory: $($dirEntry.Source) -> $($dirEntry.Target)" -ForegroundColor $ColorInfo
    } else {
        Write-Host "Warning: Data directory not found: $($dirEntry.Source)" -ForegroundColor $ColorWarning
    }
}

$QFluentResources = Get-QFluentResourceMappings -PythonExecutable $Python
foreach ($resourcePair in $QFluentResources) {
    if ($resourcePair.Length -ge 2) {
        $resSource = $resourcePair[0]
        $resTarget = $resourcePair[1]
        if (Test-Path $resSource) {
            $resolvedResSource = (Resolve-Path $resSource).Path
            $AddDataArgs += "--add-data=$resolvedResSource;$resTarget"
            Write-Host "Including qfluentwidgets resource: $resSource -> $resTarget" -ForegroundColor $ColorInfo
        }
    }
}

# Determine PyInstaller executable preference
$script:PyInstallerExe = ""
$PossibleExe = ""
try {
    $PossibleExe = Join-Path (Split-Path -Parent $Python) "pyinstaller.exe"
} catch {
    $PossibleExe = ""
}

if ($PossibleExe -and (Test-Path $PossibleExe)) {
    $script:PyInstallerExe = $PossibleExe
} elseif (Test-CommandExists "pyinstaller.exe") {
    $script:PyInstallerExe = (Get-Command "pyinstaller.exe").Source
} elseif (Test-CommandExists "pyinstaller") {
    $script:PyInstallerExe = (Get-Command "pyinstaller").Source
}

function Invoke-PyInstaller {
    param(
        [Parameter(Mandatory=$true)][string[]]$Args,
        [Parameter(Mandatory=$true)][string]$Description
    )

    Write-Host ""
    Write-Host "Running PyInstaller ($Description)..." -ForegroundColor $ColorInfo

    if ($script:PyInstallerExe) {
        Write-Host "Command: $script:PyInstallerExe $($Args -join ' ')" -ForegroundColor $ColorInfo
        & $script:PyInstallerExe @Args
    } else {
        Write-Host "Command: $Python -m PyInstaller $($Args -join ' ')" -ForegroundColor $ColorInfo
        & $Python -m PyInstaller @Args
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: PyInstaller $Description build failed" -ForegroundColor $ColorError
        exit 1
    }
}

# Compose common PyInstaller arguments
$CommonArgs = @("--noconfirm", "--clean", "--name", $AppName, "--paths", $SrcPath)

if ($Debug) {
    $CommonArgs += "--debug=all"
} else {
    $CommonArgs += "--windowed"
}

if ($NoUPX) {
    $CommonArgs += "--noupx"
    Write-Host "UPX compression disabled" -ForegroundColor $ColorWarning
}

$IconPath = $null
if (Test-Path "icon.ico") {
    $IconPath = (Resolve-Path "icon.ico").Path
    $CommonArgs += "--icon=$IconPath"
    $CommonArgs += "--add-data=$IconPath;icon.ico"
    Write-Host "Using icon.ico as application icon (PyInstaller argument & data asset)" -ForegroundColor $ColorSuccess
} else {
    Write-Host "Warning: icon.ico not found, building without custom icon" -ForegroundColor $ColorWarning
}

foreach ($module in $HiddenImports) {
    $CommonArgs += "--hidden-import=$module"
}

foreach ($module in $ExcludedModules) {
    $CommonArgs += "--exclude-module=$module"
}

foreach ($pkg in $CollectPackages) {
    $CommonArgs += "--collect-all=$pkg"
}

$CommonArgs += $AddDataArgs

Write-Host "This may take several minutes..." -ForegroundColor $ColorWarning

# Build single-file executable
$OneFileArgs = $CommonArgs + @(
    "--onefile",
    "--distpath", "dist",
    "--workpath", "build\onefile",
    "--specpath", "build\spec\onefile",
    "src\main.py"
)

Invoke-PyInstaller -Args $OneFileArgs -Description "single-file"

# Build directory distribution
$DirBuildArgs = $CommonArgs + @(
    "--onedir",
    "--distpath", "dist\${AppName}_dist",
    "--workpath", "build\onedir",
    "--specpath", "build\spec\onedir",
    "src\main.py"
)

Invoke-PyInstaller -Args $DirBuildArgs -Description "directory"

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