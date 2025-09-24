# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Import paths and dependencies
import os
import sys
from pathlib import Path

# Get the source directory
src_dir = Path('src').resolve()

def get_qfluentwidgets_resources():
    """Get qfluentwidgets resources with fallback"""
    try:
        import qfluentwidgets
        qfw_path = os.path.dirname(qfluentwidgets.__file__)
        qss_path = os.path.join(qfw_path, 'qss')
        resources_path = os.path.join(qfw_path, 'resources')

        resources = []

        # Only add if paths exist
        if os.path.exists(qss_path):
            resources.append((qss_path, 'qfluentwidgets/qss'))
        else:
            print(f"Warning: qfluentwidgets qss path not found: {qss_path}")

        if os.path.exists(resources_path):
            resources.append((resources_path, 'qfluentwidgets/resources'))
        else:
            print(f"Warning: qfluentwidgets resources path not found: {resources_path}")

        return resources
    except ImportError:
        print("Warning: qfluentwidgets not found, resources will not be included")
        return []

a = Analysis(
    ['src/main.py'],
    pathex=[str(src_dir)],
    binaries=[],
    datas=[
        # Include templates directory
        ('templates', 'templates'),
        # Include config directory
        ('config', 'config'),
        # Include database directory
        ('database', 'database'),
        # Include logs directory
        ('logs', 'logs'),
        # Include PyQt-Fluent-Widgets resources
    ] + get_qfluentwidgets_resources(),
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtNetwork',
        'PySide6.QtPrintSupport',
        'qfluentwidgets',
        'qfluentwidgets.common',
        'qfluentwidgets.components',
        'qfluentwidgets.window',
        'win32com.client',
        'win32com',
        'win32com.server',
        'pywintypes',
        'pythoncom',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
        'jinja2',
        'jinja2.ext',
        'sqlite3',
        'configparser',
        'smtplib',
        'email.mime.multipart',
        'email.mime.text',
        'email.mime.base',
        'email.encoders',
        'email.mime.message',
        'email.mime.application',
        'email.utils',
        'logging.handlers',
        'logging.config',
        'pathlib',
        'contextlib',
        'functools',
        'itertools',
        'collections',
        'collections.abc',
        'typing',
        'typing_extensions',
        'importlib',
        'importlib.util',
        'importlib.metadata',
        'pkg_resources',
        'setuptools',
        'setuptools._vendor',
        'setuptools._vendor.packaging',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PyQt5',
        'PyQt6',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'unittest',
        'test',
        'tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove duplicate entries
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path if you have one: 'assets/icon.ico'
    version=None,  # Version info file not needed for now
)

# Optional: Create a directory distribution instead of single file
# This can be useful for debugging or if single file is too slow
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