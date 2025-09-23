# Email Automation Desktop

Aplikasi desktop untuk mengirim email secara otomatis dengan monitoring folder, pengenalan file berdasarkan key/pattern, integrasi database, dan template engine untuk email massal.

## Fitur Utama

- **Folder Monitoring**: Otomatis detect file baru dengan watchdog
- **Key Extraction**: Extract key dari filename menggunakan regex pattern
- **Database Integration**: SQLite untuk data supplier dan logging
- **Email Integration**: Support Outlook (COM) dan SMTP
- **Template Engine**: Jinja2 untuk email templates
- **Modern UI**: PySide6 + PyQt-Fluent-Widgets

## Struktur Project

```
email_automation/
├── src/
│   ├── main.py                 # Entry point aplikasi
│   ├── ui/
│   │   └── main_window.py      # Main window UI
│   ├── core/
│   │   ├── folder_monitor.py   # File monitoring
│   │   ├── database_manager.py # Database operations
│   │   ├── email_sender.py     # Email client integration
│   │   ├── template_engine.py  # Template processing
│   │   └── config_manager.py   # Configuration management
│   ├── models/                 # Data models
│   └── utils/
│       └── logger.py           # Logging utility
├── database/                   # SQLite database
├── config/                     # Configuration files
├── templates/                  # Email templates
├── logs/                       # Log files
├── sent/                       # Folder untuk file terkirim
├── requirements.txt
├── email_automation.spec       # PyInstaller spec
└── build.bat                   # Build script
```

## Installation & Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Setup Database

Database akan otomatis dibuat saat pertama kali aplikasi dijalankan.

### 3. Configure Suppliers

Tambahkan data supplier ke database:

```sql
INSERT INTO suppliers (key, supplier_code, supplier_name, contact_name, emails)
VALUES ('TT003', 'TT003', 'TOKO TOKO ABADI', 'Budi Santoso', '["budi@tokoabadi.com"]');
```

### 4. Run Application

```bash
python src/main.py
```

## Building Executable

### Using PyInstaller

```bash
# Install PyInstaller
pip install pyinstaller

# Build using spec file
pyinstaller email_automation.spec

# Or use build script (Windows)
build.bat
```

### Build Output

- **Single file**: `dist/EmailAutomation.exe`
- **Directory distribution**: `dist/EmailAutomation_dist/` (recommended for faster startup)

## Configuration

### Profile Configuration

Aplikasi mendukung multiple profiles dengan konfigurasi berbeda:

```ini
[profile_invoice]
name = Invoice Orders
monitor_folder = C:/Orders/Incoming
sent_folder = C:/Orders/Sent
key_pattern = ([A-Z]{2}\d{3})
email_client = outlook
subject_template = Invoice Order - [filename_without_ext]
body_template = invoice_template.html
```

### Email Client Setup

#### Outlook
- Pastikan Outlook terinstall dan dikonfigurasi
- Aplikasi akan otomatis detect Outlook

#### SMTP (Thunderbird/Gmail/etc)
```ini
email_client = smtp
smtp_server = smtp.gmail.com
smtp_port = 587
smtp_username = your_email@gmail.com
smtp_password = your_password
smtp_use_tls = true
```

### Template Variables

Template mendukung variabel-variabel berikut:

- `{{ filename }}` - Nama file lengkap
- `{{ filename_without_ext }}` - Nama file tanpa extension
- `{{ supplier_code }}` - Kode supplier
- `{{ supplier_name }}` - Nama supplier
- `{{ contact_name }}` - Nama kontak
- `{{ date }}` - Tanggal saat ini
- `{{ time }}` - Waktu saat ini

## Usage

### 1. Setup Profile
1. Pilih atau buat profile baru
2. Set monitor folder dan sent folder
3. Konfigurasi key pattern (regex)
4. Pilih email client
5. Customize subject dan body template

### 2. Start Monitoring
1. Klik "Start Monitoring"
2. Aplikasi akan monitor folder untuk file baru
3. File yang match pattern akan otomatis diproses

### 3. Process Flow
1. File baru detected di monitor folder
2. Key extracted dari filename menggunakan regex
3. Data supplier dicari berdasarkan key
4. Email content generated menggunakan template
5. Email dikirim dengan file sebagai attachment
6. File dipindah ke sent folder
7. Log disimpan ke database

## Database Schema

### Suppliers Table
```sql
CREATE TABLE suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT UNIQUE NOT NULL,
    supplier_code TEXT NOT NULL,
    supplier_name TEXT NOT NULL,
    contact_name TEXT,
    emails TEXT NOT NULL,  -- JSON array
    cc_emails TEXT,        -- JSON array
    bcc_emails TEXT,       -- JSON array
    active BOOLEAN DEFAULT 1
);
```

### Email Logs Table
```sql
CREATE TABLE email_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path TEXT NOT NULL,
    filename TEXT NOT NULL,
    supplier_key TEXT NOT NULL,
    recipient_emails TEXT NOT NULL,
    subject TEXT NOT NULL,
    body TEXT,
    sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    status TEXT DEFAULT 'sent',
    email_client TEXT
);
```

## Troubleshooting

### Common Issues

1. **Outlook tidak detected**
   - Pastikan Outlook terinstall dan dikonfigurasi
   - Jalankan Outlook sekali untuk setup

2. **SMTP authentication error**
   - Periksa username/password
   - Untuk Gmail, gunakan App Password bukan password biasa

3. **File tidak terdetect**
   - Periksa regex pattern
   - Pastikan file extension sesuai konfigurasi

4. **Template error**
   - Periksa syntax Jinja2
   - Pastikan semua variabel tersedia

### Logging

Log tersimpan di `logs/app.log` dengan level yang bisa dikonfigurasi di `config/default.ini`.

## Development

### Tech Stack
- **Python 3.9+**
- **PySide6** - GUI framework
- **PyQt-Fluent-Widgets** - Modern UI components
- **SQLite** - Database
- **Jinja2** - Template engine
- **Watchdog** - File monitoring
- **Win32COM** - Outlook integration

### Contributing

1. Fork repository
2. Create feature branch
3. Commit changes
4. Create pull request

## License

[Add your license here]