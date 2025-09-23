# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Import paths and dependencies
import os
import sys
from pathlib import Path

# Get the source directory
src_dir = Path('src').resolve()

a = Analysis(
    ['src/main.py'],
    pathex=[str(src_dir)],
    binaries=[],
    datas=[
        # Include templates directory
        ('templates', 'templates'),
        # Include config directory
        ('config', 'config'),
        # Include PyQt-Fluent-Widgets resources
        (os.path.join(os.path.dirname(__import__('qfluentwidgets').__file__), 'qss'), 'qfluentwidgets/qss'),
        (os.path.join(os.path.dirname(__import__('qfluentwidgets').__file__), 'resources'), 'qfluentwidgets/resources'),
    ],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'qfluentwidgets',
        'win32com.client',
        'win32com',
        'pywintypes',
        'pythoncom',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
        'jinja2',
        'sqlite3',
        'configparser',
        'smtplib',
        'email.mime.multipart',
        'email.mime.text',
        'email.mime.base',
        'email.encoders',
        'logging.handlers',
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
    version='version_info.txt',  # Add version info file if needed
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