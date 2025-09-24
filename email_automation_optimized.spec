# -*- mode: python ; coding: utf-8 -*-

"""
Optimized PyInstaller spec file for Email Automation Desktop
This spec file includes only necessary dependencies and excludes unused modules.
"""

import os
import sys
from pathlib import Path

# Block cipher for additional security (None for now)
block_cipher = None

# Get the source directory
src_dir = Path('src').resolve()

def get_pyside6_fluent_resources():
    """Get PySide6-Fluent-Widgets resources"""
    try:
        import qfluentwidgets
        qfw_path = os.path.dirname(qfluentwidgets.__file__)
        
        resources = []
        resource_dirs = ['qss', 'resources', '_rc']
        
        for res_dir in resource_dirs:
            res_path = os.path.join(qfw_path, res_dir)
            if os.path.exists(res_path):
                resources.append((res_path, f'qfluentwidgets/{res_dir}'))
                print(f"Including qfluentwidgets resource: {res_path}")
            else:
                print(f"qfluentwidgets resource not found: {res_path}")
        
        return resources
    except ImportError:
        print("Warning: qfluentwidgets not found, resources will not be included")
        return []

def get_project_data_files():
    """Get project data files that need to be included"""
    data_files = []
    
    # Required directories
    required_dirs = [
        ('templates', 'templates'),
        ('config', 'config'), 
        ('database', 'database')
    ]
    
    for src_dir, dest_dir in required_dirs:
        if os.path.exists(src_dir):
            data_files.append((src_dir, dest_dir))
            print(f"Including project directory: {src_dir}")
        else:
            print(f"Warning: Project directory not found: {src_dir}")
    
    # Create logs directory in build
    logs_dir = 'logs'
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir, exist_ok=True)
    data_files.append((logs_dir, logs_dir))
    
    return data_files

# Analysis phase
a = Analysis(
    ['src/main.py'],
    pathex=[str(src_dir)],
    binaries=[],
    datas=get_project_data_files() + get_pyside6_fluent_resources(),
    hiddenimports=[
        # Python built-in modules (critical)
        'ast',
        'dis',  # Required by inspect module
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
        'darkdetect',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
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
        'readline',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove duplicate entries and optimize
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Single-file executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EmailAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Enable UPX compression by default
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Application icon
    version=None,  # Add version info if available
    uac_admin=False,  # Set to True if admin rights needed
    uac_uiaccess=False,
)

# Directory distribution (faster startup, easier debugging)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EmailAutomation_dist'
)

# Print build information
print(f"\nBuild Configuration:")
print(f"- Main script: src/main.py")
print(f"- Output name: EmailAutomation")
print(f"- Console mode: {exe.console}")
print(f"- UPX compression: {exe.upx}")
print(f"- Data files included: {len(a.datas)}")
print(f"- Hidden imports: {len(a.hiddenimports)}")
print(f"- Excluded modules: {len(a.excludes)}")
print(f"\nBuild targets:")
print(f"- Single file: dist/EmailAutomation.exe")
print(f"- Directory: dist/EmailAutomation_dist/")





