# Plan Aplikasi Email Automation Desktop

## 1. Overview Aplikasi

### Tujuan
Aplikasi desktop untuk mengirim email secara otomatis dengan monitoring folder, pengenalan file berdasarkan key/pattern, integrasi database, dan template engine untuk email massal (Invoice, Schedule Delivery, dll).

### Tech Stack
- **Language**: Python 3.9+
- **GUI Framework**: PySide6 + PyQt-Fluent-Widgets
- **Database**: SQLite
- **Template Engine**: Jinja2
- **Email Client Integration**: Outlook (win32com) / Thunderbird
- **File Monitoring**: watchdog
- **Configuration**: configparser (.ini files)

## 2. Struktur Project

```
email_automation/
├── src/
│   ├── __init__.py
│   ├── main.py                 # Entry point aplikasi
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── main_window.py      # Main window UI
│   │   ├── config_dialog.py    # Dialog konfigurasi
│   │   ├── template_editor.py  # Template editor
│   │   └── log_viewer.py       # Log viewer
│   ├── core/
│   │   ├── __init__.py
│   │   ├── folder_monitor.py   # File monitoring
│   │   ├── database_manager.py # Database operations
│   │   ├── email_sender.py     # Email client integration
│   │   ├── template_engine.py  # Template processing
│   │   ├── file_processor.py   # File processing & key extraction
│   │   └── config_manager.py   # Configuration management
│   ├── models/
│   │   ├── __init__.py
│   │   ├── supplier.py         # Data models
│   │   └── email_log.py
│   └── utils/
│       ├── __init__.py
│       ├── logger.py           # Logging utility
│       ├── app_paths.py        # Path management untuk bundling
│       └── helpers.py          # Helper functions
├── database/
│   └── email_automation.db     # SQLite database
├── config/
│   ├── default.ini             # Default configuration
│   └── profiles/               # Profile configurations
├── templates/
│   ├── invoice_template.html
│   ├── delivery_template.html
│   └── default_template.html
├── assets/
│   ├── app_icon.ico            # Application icon
│   └── images/                 # UI images
├── logs/
│   └── app.log
├── sent/                       # Folder untuk file terkirim
├── build.py                    # Build script untuk PyInstaller
├── EmailAutomation.spec        # PyInstaller spec file
├── requirements.txt
├── README.md
└── LICENSE
```

## 3. Database Schema

### Tabel: suppliers
```sql
CREATE TABLE suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT UNIQUE NOT NULL,           -- TT003
    supplier_code TEXT NOT NULL,        -- TT003 (bisa sama dengan key)
    supplier_name TEXT NOT NULL,        -- TOKO TOKO ABADI
    contact_name TEXT,                  -- Nama kontak
    emails TEXT NOT NULL,               -- JSON array: ["email1@test.com", "email2@test.com"]
    cc_emails TEXT,                     -- JSON array untuk CC
    bcc_emails TEXT,                    -- JSON array untuk BCC
    active BOOLEAN DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### Tabel: email_logs
```sql
CREATE TABLE email_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path TEXT NOT NULL,
    filename TEXT NOT NULL,
    supplier_key TEXT NOT NULL,
    recipient_emails TEXT NOT NULL,     -- JSON array
    cc_emails TEXT,
    bcc_emails TEXT,
    subject TEXT NOT NULL,
    body TEXT,
    template_used TEXT,
    sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    status TEXT DEFAULT 'sent',         -- sent, failed, pending
    error_message TEXT,
    email_client TEXT                   -- outlook, thunderbird
);
```

### Tabel: configurations
```sql
CREATE TABLE configurations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    profile_name TEXT UNIQUE NOT NULL,
    monitor_folder TEXT NOT NULL,
    sent_folder TEXT NOT NULL,
    key_pattern TEXT NOT NULL,          -- Regex pattern
    email_client TEXT NOT NULL,         -- outlook/thunderbird
    template_path TEXT,
    subject_template TEXT,
    body_template TEXT,
    active BOOLEAN DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

## 4. Core Components

### 4.1 Folder Monitor (folder_monitor.py)
```python
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import os

class FileHandler(FileSystemEventHandler):
    def __init__(self, callback, key_pattern):
        self.callback = callback
        self.key_pattern = re.compile(key_pattern)
    
    def on_created(self, event):
        if not event.is_directory:
            self.process_file(event.src_path)
    
    def process_file(self, file_path):
        filename = os.path.basename(file_path)
        match = self.key_pattern.search(filename)
        if match:
            key = match.group(1)  # Ambil group pertama dari regex
            self.callback(file_path, key)

class FolderMonitor:
    def __init__(self):
        self.observer = Observer()
        self.is_monitoring = False
    
    def start_monitoring(self, folder_path, callback, key_pattern):
        handler = FileHandler(callback, key_pattern)
        self.observer.schedule(handler, folder_path, recursive=False)
        self.observer.start()
        self.is_monitoring = True
    
    def stop_monitoring(self):
        if self.observer.is_alive():
            self.observer.stop()
            self.observer.join()
        self.is_monitoring = False
```

### 4.2 Database Manager (database_manager.py)
```python
import sqlite3
import json
from contextlib import contextmanager

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_database()
    
    @contextmanager
    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()
    
    def get_supplier_by_key(self, key):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM suppliers WHERE key = ? AND active = 1", (key,))
            row = cursor.fetchone()
            if row:
                return {
                    'id': row['id'],
                    'key': row['key'],
                    'supplier_code': row['supplier_code'],
                    'supplier_name': row['supplier_name'],
                    'contact_name': row['contact_name'],
                    'emails': json.loads(row['emails']),
                    'cc_emails': json.loads(row['cc_emails'] or '[]'),
                    'bcc_emails': json.loads(row['bcc_emails'] or '[]')
                }
        return None
    
    def log_email_sent(self, log_data):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO email_logs 
                (file_path, filename, supplier_key, recipient_emails, cc_emails, 
                 bcc_emails, subject, body, template_used, email_client, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                log_data['file_path'], log_data['filename'], log_data['supplier_key'],
                json.dumps(log_data['recipient_emails']), 
                json.dumps(log_data.get('cc_emails', [])),
                json.dumps(log_data.get('bcc_emails', [])),
                log_data['subject'], log_data['body'], log_data.get('template_used'),
                log_data['email_client'], log_data.get('status', 'sent')
            ))
            conn.commit()
```

### 4.3 Email Sender (email_sender.py)
```python
import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

class OutlookSender:
    def __init__(self):
        try:
            self.outlook = win32.Dispatch('outlook.application')
        except:
            raise Exception("Outlook tidak terinstall atau tidak dapat diakses")
    
    def send_email(self, to_emails, cc_emails, bcc_emails, subject, body, attachment_path):
        mail = self.outlook.CreateItem(0)  # 0 = olMailItem
        
        mail.To = '; '.join(to_emails)
        if cc_emails:
            mail.CC = '; '.join(cc_emails)
        if bcc_emails:
            mail.BCC = '; '.join(bcc_emails)
        
        mail.Subject = subject
        mail.HTMLBody = body
        
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)
        
        mail.Send()
        return True

class ThunderbirdSender:
    def __init__(self, smtp_server, smtp_port, username, password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.username = username
        self.password = password
    
    def send_email(self, to_emails, cc_emails, bcc_emails, subject, body, attachment_path):
        msg = MIMEMultipart()
        msg['From'] = self.username
        msg['To'] = ', '.join(to_emails)
        if cc_emails:
            msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {os.path.basename(attachment_path)}'
                )
                msg.attach(part)
        
        server = smtplib.SMTP(self.smtp_server, self.smtp_port)
        server.starttls()
        server.login(self.username, self.password)
        
        all_recipients = to_emails + (cc_emails or []) + (bcc_emails or [])
        server.sendmail(self.username, all_recipients, msg.as_string())
        server.quit()
        return True
```

### 4.4 Template Engine (template_engine.py)
```python
from jinja2 import Template, Environment, FileSystemLoader
import os
import re

class EmailTemplateEngine:
    def __init__(self, template_dir):
        self.env = Environment(loader=FileSystemLoader(template_dir))
    
    def render_template(self, template_content, variables):
        """Render template dengan variabel yang diberikan"""
        template = Template(template_content)
        return template.render(**variables)
    
    def prepare_variables(self, file_path, supplier_data, custom_vars=None):
        """Siapkan variabel untuk template"""
        filename = os.path.basename(file_path)
        filename_without_ext = os.path.splitext(filename)[0]
        
        variables = {
            'filename': filename,
            'filename_without_ext': filename_without_ext,
            'filepath': file_path,
            'supplier_code': supplier_data.get('supplier_code', ''),
            'supplier_name': supplier_data.get('supplier_name', ''),
            'contact_name': supplier_data.get('contact_name', ''),
            'emails': supplier_data.get('emails', []),
            'cc_emails': supplier_data.get('cc_emails', []),
            'bcc_emails': supplier_data.get('bcc_emails', []),
        }
        
        # Tambahkan custom variables jika ada
        if custom_vars:
            variables.update(custom_vars)
        
        return variables
    
    def process_simple_variables(self, text, variables):
        """Process variabel sederhana [variable_name]"""
        def replace_var(match):
            var_name = match.group(1)
            return str(variables.get(var_name, f'[{var_name}]'))
        
        return re.sub(r'\[(\w+)\]', replace_var, text)
```

## 5. User Interface Components

### 5.1 Main Window Layout
- **Menu Bar**: File, Settings, View, Help
- **Toolbar**: Start/Stop Monitoring, Configuration, Template Editor
- **Main Area**:
  - Configuration Panel (kiri)
    - Profile Selection
    - Monitor Folder Path
    - Sent Folder Path
    - Key Pattern (Regex)
    - Email Client Selection
  - Template Panel (tengah)
    - Subject Template
    - Body Template
    - Preview Area
  - Status Panel (kanan)
    - Monitoring Status
    - Recent Files
    - Email Queue
- **Bottom Panel**: Log viewer dengan filtering

### 5.2 Configuration Dialog
- **Profile Management**
- **Database Connection Settings**
- **Email Client Configuration**
  - Outlook: Auto-detect
  - Thunderbird: SMTP Settings
- **Default Templates**
- **File Patterns**

### 5.3 Template Editor
- **Rich Text Editor** untuk body template
- **Variable Helper Panel**
- **Preview Panel** dengan sample data
- **Template Validation**

## 6. Configuration File Structure (.ini)

```ini
[DEFAULT]
current_profile = profile1
database_path = database/email_automation.db
log_level = INFO

[profile1]
name = Invoice Orders
monitor_folder = C:/Orders/Incoming
sent_folder = C:/Orders/Sent
key_pattern = ([A-Z]{2}\d{3})
email_client = outlook
subject_template = Invoice Order - [filename_without_ext]
body_template = templates/invoice_template.html
auto_start = true
file_extensions = .pdf,.xlsx,.docx

[profile2]
name = Delivery Schedule
monitor_folder = C:/Delivery/Incoming
sent_folder = C:/Delivery/Sent
key_pattern = DELIVERY_([A-Z0-9]+)
email_client = thunderbird
smtp_server = smtp.gmail.com
smtp_port = 587
smtp_username = your_email@gmail.com
smtp_password = your_password
```

## 7. Implementation Phases

### Phase 1: Core Infrastructure (Week 1-2)
- Setup project structure
- Database schema implementation
- Basic configuration management
- File monitoring system

### Phase 2: Email Integration (Week 2-3)
- Outlook integration
- Thunderbird integration
- Template engine implementation
- Basic email sending functionality

### Phase 3: User Interface (Week 3-4)
- Main window layout
- Configuration dialogs
- Template editor
- Log viewer

### Phase 4: Advanced Features (Week 4-5)
- Profile management
- Advanced template features
- Error handling and retry logic
- Performance optimization

### Phase 5: Testing & Polish (Week 5-6)
- Unit testing
- Integration testing
- UI/UX improvements
- Documentation

### Phase 6: Bundling & Distribution (Week 6)
- PyInstaller configuration
- Build script automation
- Portable deployment testing
- Installation package creation
- Code signing (optional)

## 8. Key Features Implementation

### 8.1 File Processing Flow
1. **File Detection**: Watchdog mendeteksi file baru
2. **Key Extraction**: Regex pattern extract key dari filename
3. **Database Lookup**: Cari supplier berdasarkan key
4. **Template Processing**: Generate email content
5. **Email Sending**: Send via Outlook/Thunderbird
6. **File Management**: Move ke sent folder
7. **Logging**: Record ke database dan log file

### 8.2 Template Variables
- `[filename]`: Nama file lengkap
- `[filename_without_ext]`: Nama file tanpa extension
- `[supplier_code]`: Kode supplier
- `[supplier_name]`: Nama supplier
- `[contact_name]`: Nama kontak
- `[emails]`: Email addresses (array)
- `[date]`: Tanggal saat ini
- `[time]`: Waktu saat ini
- Custom variables dari database

### 8.3 Error Handling
- File access errors
- Database connection errors
- Email client errors
- Network errors
- Template rendering errors
- Configuration errors

## 9. Dependencies (requirements.txt)

```txt
PySide6==6.5.0
PyQt-Fluent-Widgets==1.4.0
watchdog==3.0.0
Jinja2==3.1.2
pywin32==306  # for Outlook integration
pyinstaller==5.13.2  # for bundling
sqlite3  # built-in
configparser  # built-in
logging  # built-in
re  # built-in
os  # built-in
json  # built-in
```

## 9.1. Application Bundling dengan PyInstaller

### 9.1.1 Build Script (build.py)
```python
import PyInstaller.__main__
import os
import shutil
import sys

def build_application():
    # Clean previous build
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    # PyInstaller arguments
    args = [
        'src/main.py',
        '--name=EmailAutomation',
        '--onedir',  # atau --onefile untuk single executable
        '--windowed',  # Hide console window
        '--icon=assets/app_icon.ico',
        '--add-data=templates;templates',
        '--add-data=config;config',
        '--add-data=database;database',
        '--add-data=assets;assets',
        '--hidden-import=win32com.client',
        '--hidden-import=pywintypes',
        '--hidden-import=pythoncom',
        '--hidden-import=sqlite3',
        '--collect-all=PyQt-Fluent-Widgets',
        '--collect-all=PySide6',
        '--collect-all=jinja2',
        '--collect-all=watchdog',
        '--exclude-module=matplotlib',
        '--exclude-module=numpy',
        '--exclude-module=pandas',
        '--clean',
    ]
    
    PyInstaller.__main__.run(args)
    
    # Copy additional files
    post_build_setup()

def post_build_setup():
    """Setup files setelah build"""
    dist_dir = 'dist/EmailAutomation'
    
    # Create necessary directories
    os.makedirs(f'{dist_dir}/logs', exist_ok=True)
    os.makedirs(f'{dist_dir}/sent', exist_ok=True)
    os.makedirs(f'{dist_dir}/config/profiles', exist_ok=True)
    
    # Copy default database if not exists
    if not os.path.exists(f'{dist_dir}/database/email_automation.db'):
        # Create empty database with schema
        create_default_database(f'{dist_dir}/database/email_automation.db')
    
    print("Build completed successfully!")
    print(f"Executable location: {dist_dir}/EmailAutomation.exe")

def create_default_database(db_path):
    """Create default database dengan schema"""
    import sqlite3
    
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Create tables
    cursor.execute('''
        CREATE TABLE suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key TEXT UNIQUE NOT NULL,
            supplier_code TEXT NOT NULL,
            supplier_name TEXT NOT NULL,
            contact_name TEXT,
            emails TEXT NOT NULL,
            cc_emails TEXT,
            bcc_emails TEXT,
            active BOOLEAN DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE email_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_path TEXT NOT NULL,
            filename TEXT NOT NULL,
            supplier_key TEXT NOT NULL,
            recipient_emails TEXT NOT NULL,
            cc_emails TEXT,
            bcc_emails TEXT,
            subject TEXT NOT NULL,
            body TEXT,
            template_used TEXT,
            sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'sent',
            error_message TEXT,
            email_client TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE configurations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            profile_name TEXT UNIQUE NOT NULL,
            monitor_folder TEXT NOT NULL,
            sent_folder TEXT NOT NULL,
            key_pattern TEXT NOT NULL,
            email_client TEXT NOT NULL,
            template_path TEXT,
            subject_template TEXT,
            body_template TEXT,
            active BOOLEAN DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Insert sample data
    cursor.execute('''
        INSERT INTO suppliers (key, supplier_code, supplier_name, contact_name, emails)
        VALUES (?, ?, ?, ?, ?)
    ''', ('TT003', 'TT003', 'TOKO TOKO ABADI', 'John Doe', '["john@tokotokoabadi.com", "admin@tokotokoabadi.com"]'))
    
    conn.commit()
    conn.close()

if __name__ == '__main__':
    build_application()
```

### 9.1.2 PyInstaller Spec File (EmailAutomation.spec)
```python
# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

# Collect all required packages
datas = []
binaries = []
hiddenimports = []

# PyQt-Fluent-Widgets
tmp_ret = collect_all('qfluentwidgets')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# PySide6
tmp_ret = collect_all('PySide6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Add application data
datas += [
    ('templates', 'templates'),
    ('config', 'config'),
    ('database', 'database'),
    ('assets', 'assets'),
]

# Hidden imports
hiddenimports += [
    'win32com.client',
    'pywintypes',
    'pythoncom',
    'sqlite3',
    'jinja2',
    'watchdog.observers',
    'watchdog.events',
]

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'PIL',
        'tkinter', 'test', 'unittest', 'pydoc'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EmailAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Hide console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/app_icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EmailAutomation'
)
```

### 9.1.3 Portable Application Structure
```
EmailAutomation/ (Distributed folder)
├── EmailAutomation.exe          # Main executable
├── _internal/                   # PyInstaller internals
├── config/
│   ├── default.ini             # Default configuration
│   └── profiles/               # User profiles
├── database/
│   └── email_automation.db     # SQLite database
├── templates/
│   ├── invoice_template.html
│   ├── delivery_template.html
│   └── default_template.html
├── logs/                       # Application logs
├── sent/                       # Default sent folder
├── assets/
│   ├── app_icon.ico
│   └── images/
├── README.txt                  # User manual
└── uninstall.bat              # Uninstaller (optional)
```

### 9.1.4 Application Path Management (utils/app_paths.py)
```python
import os
import sys
from pathlib import Path

class AppPaths:
    """Manage application paths untuk bundled dan development mode"""
    
    def __init__(self):
        self._base_path = self._get_base_path()
        self._ensure_directories()
    
    def _get_base_path(self):
        """Get base path aplikasi"""
        if getattr(sys, 'frozen', False):
            # Running as bundled executable
            return Path(sys.executable).parent
        else:
            # Running in development
            return Path(__file__).parent.parent
    
    def _ensure_directories(self):
        """Ensure required directories exist"""
        directories = [
            'logs', 'sent', 'config/profiles', 
            'database', 'templates'
        ]
        
        for directory in directories:
            (self._base_path / directory).mkdir(parents=True, exist_ok=True)
    
    @property
    def base_path(self):
        return self._base_path
    
    @property
    def config_path(self):
        return self._base_path / 'config'
    
    @property
    def database_path(self):
        return self._base_path / 'database' / 'email_automation.db'
    
    @property
    def templates_path(self):
        return self._base_path / 'templates'
    
    @property
    def logs_path(self):
        return self._base_path / 'logs'
    
    @property
    def sent_path(self):
        return self._base_path / 'sent'
    
    def get_config_file(self, filename='default.ini'):
        return self.config_path / filename
    
    def get_template_file(self, filename):
        return self.templates_path / filename
    
    def get_profile_config(self, profile_name):
        return self.config_path / 'profiles' / f'{profile_name}.ini'

# Global instance
app_paths = AppPaths()
```

### 9.1.5 Build Commands

#### Development Build
```bash
# Install dependencies
pip install -r requirements.txt

# Build application
python build.py

# atau manual dengan spec file
pyinstaller EmailAutomation.spec
```

#### Production Build
```bash
# Create build environment
python -m venv build_env
build_env\Scripts\activate
pip install -r requirements.txt

# Build
python build.py

# Test executable
dist\EmailAutomation\EmailAutomation.exe
```

### 9.1.6 Configuration untuk Portable Mode (config_manager.py)
```python
import configparser
import os
from utils.app_paths import app_paths

class ConfigManager:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_file = app_paths.get_config_file()
        self.load_config()
    
    def load_config(self):
        """Load configuration dari file atau create default"""
        if not self.config_file.exists():
            self.create_default_config()
        
        self.config.read(self.config_file, encoding='utf-8')
    
    def create_default_config(self):
        """Create default configuration"""
        self.config['DEFAULT'] = {
            'current_profile': 'default',
            'database_path': str(app_paths.database_path),
            'log_level': 'INFO'
        }
        
        self.config['default'] = {
            'name': 'Default Profile',
            'monitor_folder': str(app_paths.base_path / 'incoming'),
            'sent_folder': str(app_paths.sent_path),
            'key_pattern': r'([A-Z]{2}\d{3})',
            'email_client': 'outlook',
            'subject_template': 'Document - [filename_without_ext]',
            'body_template': 'default_template.html'
        }
        
        self.save_config()
    
    def save_config(self):
        """Save configuration ke file"""
        app_paths.config_path.mkdir(exist_ok=True)
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
```

### 9.1.7 Deployment Considerations

#### Single File vs Directory
- **Directory Mode** (recommended): Faster startup, easier debugging
- **Single File Mode**: Portable tapi slower startup

#### Size Optimization
- Exclude unnecessary modules
- Use UPX compression
- Remove debug symbols

#### Security Considerations
- Code signing untuk avoid antivirus false positives
- Obfuscation jika diperlukan

#### Auto-updater (Optional)
```python
class AutoUpdater:
    def __init__(self):
        self.update_server = "https://your-server.com/updates/"
        self.current_version = "1.0.0"
    
    def check_for_updates(self):
        # Implementation untuk check updates
        pass
    
    def download_update(self):
        # Implementation untuk download update
        pass
```

## 10. Testing Strategy

### Unit Tests
- Database operations
- Template engine
- File processing
- Configuration management

### Integration Tests
- Email client integration
- File monitoring
- End-to-end workflow

### UI Tests
- User interactions
- Dialog functionality
- Data validation