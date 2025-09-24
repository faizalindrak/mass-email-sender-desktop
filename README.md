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
---

# Panduan Penggunaan Lengkap

Dokumentasi ini menjelaskan langkah-langkah lengkap untuk menginstal, mengonfigurasi, dan menggunakan aplikasi Email Automation Desktop, termasuk detail antarmuka, pengaturan profile, template, integrasi email client, monitoring folder, logging, serta troubleshooting umum.

Referensi komponen penting di dalam kode:
- Entry point aplikasi: [def main()](src/main.py:38)
- Jendela utama UI: [class MainWindow](src/ui/main_window.py:106)
- Worker otomasi email: [class EmailAutomationWorker](src/ui/main_window.py:22)
- Monitoring folder: [class FolderMonitor](src/core/folder_monitor.py:47), handler file [class FileHandler](src/core/folder_monitor.py:9)
- Manajemen konfigurasi: [class ConfigManager](src/core/config_manager.py:6), ambil profile [def get_profile_config()](src/core/config_manager.py:69)
- Engine template: [class EmailTemplateEngine](src/core/template_engine.py:7), proses variabel sederhana [def process_simple_variables()](src/core/template_engine.py:79), membuat template default [def create_default_templates()](src/core/template_engine.py:116)
- Pengirim email: [class EmailSenderFactory](src/core/email_sender.py:148), Outlook [class OutlookSender](src/core/email_sender.py:21), SMTP/Thunderbird [class ThunderbirdSender](src/core/email_sender.py:73)

## 1. Prasyarat Sistem

- Sistem operasi: Windows 10/11 (disarankan Windows 11)
- Python: 3.9 atau lebih tinggi
- Outlook (untuk integrasi COM) jika menggunakan client `outlook`
- Akses SMTP (mis. Gmail/Thunderbird/Office 365) jika menggunakan client `smtp`/`thunderbird`
- Hak akses baca/tulis ke folder yang akan dimonitor dan folder `sent`

## 2. Instalasi

Anda dapat menggunakan skrip setup untuk menyiapkan lingkungan virtual dan dependensi.

Opsi A — Menggunakan skrip setup:
```bat
setup.bat
```
- Membuat virtual environment `venv`
- Menginstal dependensi dari `requirements.txt`

Opsi B — Manual:
```bash
pip install -r requirements.txt
```

Validasi cepat:
```bash
python -V
pip -V
```

Jika modul GUI gagal diimpor, pesan akan muncul dari fungsi [def main()](src/main.py:38). Pastikan dependensi berikut terpasang:
```bash
pip install PySide6 PyQt-Fluent-Widgets watchdog Jinja2 pywin32
```

## 3. Menjalankan Aplikasi

Jalankan entry point:
```bash
python src/main.py
```

Apa yang terjadi saat startup:
- Inisialisasi Qt dan tema Fluent (lihat [def main()](src/main.py:38))
- Membaca konfigurasi melalui [class ConfigManager](src/core/config_manager.py:6)
- Menginisialisasi logging (file log default di `logs/app.log`)
- Membuat template default bila belum ada dengan [def create_default_templates()](src/core/template_engine.py:116)
- Menampilkan jendela utama [class MainWindow](src/ui/main_window.py:106)

## 4. Struktur Antarmuka (UI) dan Alur Kerja

Antarmuka utama (lihat [class MainWindow](src/ui/main_window.py:106)) terdiri dari tiga panel:

1) Panel Kiri — Configuration
- Profile: pilih profile aktif, load/save profile ke/dari JSON
- Database File: path database SQLite
- Monitor Folder: folder sumber yang dipantau
- Sent Folder: folder tujuan perpindahan file setelah email terkirim
- Key Pattern (Regex): pola untuk ekstraksi key dari nama file
- Email Client: pilih `outlook`, `thunderbird`, atau `smtp`
- Constant Variables: default CC/BCC + 2 variabel custom

2) Panel Tengah — Template & Preview
- Email Form tab: To, CC, BCC, Subject, Template file, dan Body editor
- Variables tab: daftar variabel yang tersedia dan sample data preview
- Preview tab: tombol “Generate Preview” menghasilkan tampilan HTML dari subject/body yang sudah diproses variabel sederhana

3) Panel Kanan — Status & Logs
- Status monitoring (Active/Stopped) dan jumlah file yang diproses
- Recent Files: daftar file terbaru yang terdeteksi, termasuk keberhasilan/gagal
- Email Logs: tabel log email (Time, File, Supplier, Status)

Tombol-tombol utama:
- Start Monitoring: [def start_monitoring()](src/ui/main_window.py:553)
- Stop Monitoring: [def stop_monitoring()](src/ui/main_window.py:584)
- Send Test Email: [def send_test_email()](src/ui/main_window.py:757)
- Generate Preview: [def generate_preview()](src/ui/main_window.py:907)
- Load/Save Profile dari/ke file JSON: [def load_profile_from_file()](src/ui/main_window.py:667), [def save_profile_to_file()](src/ui/main_window.py:695)

## 5. Konfigurasi

Konfigurasi global berada di `config/default.ini` dan dikelola oleh [class ConfigManager](src/core/config_manager.py:6). Contoh:

```ini
[DEFAULT]
current_profile = default
database_path = database/email_automation.db
log_level = INFO
log_file = logs/app.log
template_dir = templates
auto_start_monitoring = false

[profile_default]
name = Default Profile
monitor_folder =
sent_folder = sent
key_pattern = ([A-Z]{2}\d{3})
email_client = outlook
subject_template = Document - [filename_without_ext]
body_template = default_template.html
auto_start = false
file_extensions = .pdf,.xlsx,.docx,.txt
```

Catatan penting:
- `current_profile`: nama profile aktif (tanpa prefix `profile_`)
- `file_extensions`: daftar ekstensi diizinkan; file di luar daftar akan di-skip oleh [class FileHandler](src/core/folder_monitor.py:9)
- `key_pattern`: gunakan grup tangkap (capture group) jika ingin mengambil bagian tertentu, misal `([A-Z]{2}\d{3})` akan mengambil `TT003` dari `TT003_invoice_2024.pdf`

### 5.1. Konfigurasi Profile Tambahan

Contoh profile Invoice:
```ini
[profile_invoice]
name = Invoice Orders
monitor_folder = C:/Orders/Incoming
sent_folder = C:/Orders/Sent
key_pattern = ([A-Z]{2}\d{3})
email_client = outlook
subject_template = Invoice Order - [filename_without_ext]
body_template = invoice_template.html
file_extensions = .pdf,.xlsx,.docx
```

Contoh profile Delivery menggunakan SMTP:
```ini
[profile_delivery]
name = Delivery Schedule
monitor_folder = C:/Delivery/Incoming
sent_folder = C:/Delivery/Sent
key_pattern = DELIVERY_([A-Z0-9]+)
email_client = smtp
smtp_server = smtp.gmail.com
smtp_port = 587
smtp_username = your_email@gmail.com
smtp_password = your_app_password
smtp_use_tls = true
subject_template = Delivery Schedule - [filename_without_ext]
body_template = delivery_template.html
file_extensions = .pdf,.xlsx
```

Validasi profile dilakukan oleh [def validate_profile_config()](src/core/config_manager.py:167). Pastikan `monitor_folder` ada, `email_client` valid, dan bila SMTP maka field `smtp_*` terisi.

### 5.2. Manajemen Profile via UI

- Load Profile dari JSON: [def load_profile_from_file()](src/ui/main_window.py:667)
- Save Profile ke JSON: [def save_profile_to_file()](src/ui/main_window.py:695)
- Export/Import via API ConfigManager:
  - Export: [def export_profile()](src/core/config_manager.py:200)
  - Import: [def import_profile()](src/core/config_manager.py:206)

## 6. Database

Skema penting:
- Tabel `suppliers`: menyimpan data supplier, termasuk daftar email (JSON)
- Tabel `email_logs`: menyimpan riwayat pengiriman email

Pengisian contoh supplier:
```sql
INSERT INTO suppliers (key, supplier_code, supplier_name, contact_name, emails)
VALUES ('TT003', 'TT003', 'TOKO TOKO ABADI', 'Budi Santoso', '["budi@tokoabadi.com"]');
```

Lookup supplier berdasarkan key dilakukan di DatabaseManager (lihat file `src/core/database_manager.py`). UI menampilkan log melalui pemanggilan `get_email_logs(...)` dari DatabaseManager (tabel di panel kanan).

Tips:
- Gunakan DB Browser for SQLite untuk melihat/mengedit isi database `database/email_automation.db`
- Pastikan field `emails` berisi JSON array string yang valid (mis. `["a@b.com","c@d.com"]`)

## 7. Template Email

Direktori template: `templates/` (dikendalikan oleh [def get_template_dir()](src/core/config_manager.py:153))

Template default akan dibuat otomatis oleh [def create_default_templates()](src/core/template_engine.py:116) bila belum ada:
- `invoice_template.html`
- `delivery_template.html`
- `default_template.html`

### 7.1. Variabel di Template

Engine mendukung dua cara:
1) Jinja2 (di file `.html`)
   - Contoh: `{{ filename_without_ext }}`, `{{ supplier_name }}`, `{{ date }}`
2) Variabel sederhana dengan format `[variable_name]` di Subject/Body editor UI (diproses oleh [def process_simple_variables()](src/core/template_engine.py:79))
   - Contoh: Subject: `Invoice - [filename_without_ext]`

Variabel umum yang tersedia (lihat [def prepare_variables()](src/core/template_engine.py:33)):
- File: `filename`, `filename_without_ext`, `filepath`, `file_size`, `file_size_mb`
- Supplier: `supplier_code`, `supplier_name`, `contact_name`, `emails`, `cc_emails`, `bcc_emails`
- Waktu: `date`, `time`, `datetime`, `date_indonesian`, `day`, `month`, `year`, `month_name`, `day_name`
- Sistem: `current_user`, `computer_name`

### 7.2. Menggunakan Template di UI

- Pilih file template pada combobox “Template” di tab Email Form
- Konten template akan dimuat ke editor “Body”
- Klik “Generate Preview” untuk melihat hasil rendering gabungan variabel sederhana (subject/body) dengan sample data

### 7.3. Panduan Penggunaan Variabel di Email Form

Bagian Email Form mendukung placeholder variabel sederhana dengan format [variable_name] pada field Subject dan Body. Evaluasi placeholder dilakukan oleh fungsi [def process_simple_variables()](src/core/template_engine.py:79).

- Cara memasukkan variabel dari UI:
  - Buka tab “Variables”
  - Pilih variabel dari daftar “Available Variables”
  - Klik “Insert to Subject” atau “Insert to Body”
  - Tombol tersebut memanggil [def insert_variable_to_subject()](src/ui/main_window.py:882) dan [def insert_variable_to_body()](src/ui/main_window.py:896)

- Aturan penulisan variabel sederhana:
  - Format: [nama_variabel] (huruf/angka/underscore saja — sesuai pola \w+)
  - Jika variabel tidak ditemukan saat proses, teks akan dibiarkan apa adanya, misal [unknown_var]
  - Jika nilai variabel berupa list (contoh: emails, cc_emails, bcc_emails), akan digabung menjadi string dipisah koma dan spasi, contoh: "a@b.com, c@d.com"
  - Variabel yang tersedia untuk placeholder sederhana merujuk pada data yang disiapkan oleh [def prepare_variables()](src/core/template_engine.py:33), antara lain:
    - File: [filename], [filename_without_ext], [filepath], [file_size], [file_size_mb]
    - Supplier: [supplier_code], [supplier_name], [contact_name], [emails], [cc_emails], [bcc_emails]
    - Waktu: [date], [time], [datetime], [date_indonesian], [day], [month], [year], [month_name], [day_name]
    - Sistem: [current_user], [computer_name]
  - Variabel custom dari panel “Constant Variables”:
    - Isi “Custom Variable 1/2” (name dan value). Selama “name” hanya berisi huruf/angka/underscore, Anda dapat memakainya sebagai [name] di Subject/Body
    - Catatan: variabel custom ini digunakan dalam Preview. Alur otomatis menggunakan konfigurasi profil dan data supplier

- Perilaku di Preview vs. Kirim Email:
  - Preview:
    - Tombol “Generate Preview” akan mengganti placeholder [var] pada Subject/Body menggunakan [def generate_preview()](src/ui/main_window.py:907)
  - Send Test Email (Manual):
    - Tombol “Send Test Email” mengirim apa yang tertulis di Subject/Body tanpa penggantian placeholder [def send_test_email()](src/ui/main_window.py:757)
  - Alur Otomatis (Monitoring):
    - Subject diambil dari profile (subject_template) dan diproses dengan placeholder sederhana oleh [def process_file()](src/ui/main_window.py:35)
    - Body di-render dari file template (Jinja2) melalui [def render_file_template()](src/core/template_engine.py:25), bukan placeholder [var]

- Menggabungkan dengan Jinja2:
  - Di file template HTML (mis. default_template.html), gunakan sintaks Jinja2 seperti {{ supplier_name }} untuk body otomatis
  - Placeholder [var] cocok untuk Subject dan editor Body di Email Form saat Preview. Untuk file template .html, gunakan Jinja2, bukan [var]

- Tips dan catatan:
  - Menampilkan tanda kurung siku literal:
    - Placeholder dikenali hanya jika berisi huruf/angka/underscore. Untuk menampilkan teks “[filename]” secara literal, sisipkan karakter non-word, misalnya “[file name]” atau “[file-name]”
  - Hindari tanda kurung ganda bersarang, seperti [[filename]], karena tidak didukung dan dapat menghasilkan sisa karakter bracket
  - Nama variabel custom hanya boleh berisi huruf/angka/underscore. Tanda hubung (dash) tidak cocok untuk placeholder sederhana
  - Gunakan daftar “Available Variables” untuk mengurangi salah ketik saat memasukkan placeholder

- Contoh penggunaan:
  - Subject:
    - “Invoice - [filename_without_ext] untuk [supplier_name]”
  - Body (Email Form editor):
    - “Dear [contact_name], mohon review lampiran [filename] untuk supplier [supplier_code].”
  - Body (Template HTML Jinja2):
    - “Dear {{ contact_name }}, mohon review lampiran {{ filename }} untuk supplier {{ supplier_code }}.”

## 8. Monitoring Folder dan Alur Otomasi

Monitoring dikendalikan oleh [class FolderMonitor](src/core/folder_monitor.py:47):

- Mulai monitoring: [def start_monitoring()](src/core/folder_monitor.py:55)
  - Auto membuat subfolder `sent` jika belum ada
  - Menjadwalkan handler [class FileHandler](src/core/folder_monitor.py:9)
  - Menyaring file berdasarkan `file_extensions`
  - Mengekstrak key berdasarkan `key_pattern`
- Hentikan monitoring: [def stop_monitoring()](src/core/folder_monitor.py:86)
- Pindahkan file ke sent: [def move_file_to_sent()](src/core/folder_monitor.py:134)

Alur kerja pemrosesan file (lihat [class EmailAutomationWorker](src/ui/main_window.py:22) → [def process_file()](src/ui/main_window.py:35)):
1) File terdeteksi di folder monitor
2) Ekstraksi `key` dari nama file (regex)
3) Ambil data supplier berdasar `key`
4) Siapkan variabel template
5) Render Subject (variabel sederhana) dan Body (template file Jinja2)
6) Buat sender berdasarkan `email_client` via [class EmailSenderFactory](src/core/email_sender.py:148)
7) Kirim email dengan attachment file
8) Logging ke database
9) Pindahkan file ke folder `sent`

## 9. Pengaturan Email Client

Factory memilih pengirim berdasarkan `email_client`:
- `outlook`: [class OutlookSender](src/core/email_sender.py:21), memerlukan Outlook terpasang dan profil mail tersetup (COM automation)
- `smtp`/`thunderbird`: [class ThunderbirdSender](src/core/email_sender.py:73), memerlukan konfigurasi `smtp_server`, `smtp_port`, `smtp_username`, `smtp_password`, `smtp_use_tls`

Catatan khusus Gmail:
- Aktifkan 2FA
- Buat “App Password”
- Gunakan `smtp.gmail.com:587` dengan TLS

## 10. Langkah Penggunaan (Step-by-step)

1) Siapkan profile dan database
- Edit `config/default.ini` atau buat profile baru di UI
- Masukkan supplier minimal satu data dengan email valid

2) Buka aplikasi
```bash
python src/main.py
```

3) Di UI (panel kiri):
- Pilih Profile (atau load dari file JSON)
- Pastikan `Monitor Folder` menunjuk ke folder yang ingin dipantau
- `Sent Folder` di-set (default akan otomatis dibuat sebagai subfolder `sent`)
- Set `Key Pattern` sesuai format nama file
- Pilih `Email Client` (Outlook/SMTP)
- Isi Default CC/BCC bila diperlukan

4) Di UI (panel tengah):
- Pilih template (invoice/delivery/default atau kustom)
- Isi Subject dan Body (gunakan variabel sederhana `[filename_without_ext]`, dsb.)
- Klik “Generate Preview” untuk verifikasi tampilan

5) Start monitoring:
- Klik “Start Monitoring”
- Drop file baru ke folder monitor sesuai pola; aplikasi akan otomatis memproses dan mengirim email

6) Lihat hasil:
- Panel Status menunjukkan jumlah file diproses
- Panel Recent Files menampilkan keberhasilan/gagal
- Panel Email Logs menampilkan riwayat (waktu, file, supplier, status)

7) Stop monitoring:
- Klik “Stop Monitoring”

8) Kirim email uji (opsional):
- Gunakan “Send Test Email” (tidak ada lampiran)
- Pastikan konfigurasi SMTP/Outlook sudah benar

## 11. Logging

- File log: `logs/app.log`
- Level log diatur oleh `log_level` dalam `config/default.ini`
- Proses monitoring dan pengiriman email akan menulis informasi ke log (lihat [class FolderMonitor](src/core/folder_monitor.py:47) dan sender di `src/core/email_sender.py`)

## 12. Build dan Distribusi

Bangun executable (Windows) menggunakan PyInstaller dan spesifikasi `email_automation.spec`:
```bat
build.bat
```
Atau manual:
```bash
pip install pyinstaller
pyinstaller email_automation.spec
```

Output:
- Single file: `dist/EmailAutomation.exe`
- Directory distribution (start lebih cepat): `dist/EmailAutomation_dist/EmailAutomation.exe`

## 13. Troubleshooting

- Outlook tidak terdeteksi
  - Pastikan Outlook terinstal dan sudah dijalankan minimal sekali
  - Jalankan aplikasi sebagai Administrator jika diperlukan

- SMTP authentication error
  - Periksa `smtp_server`, `smtp_port`, `smtp_username`, `smtp_password`, `smtp_use_tls`
  - Gunakan App Password (khusus Gmail/Office 365 jika kebijakan keamanan ketat)

- File tidak terdeteksi
  - Pastikan `file_extensions` mengizinkan ekstensi file
  - Verifikasi `key_pattern` sesuai format nama file (gunakan grup tangkap jika perlu)
  - Cek hak akses folder monitor

- Template error
  - Validasi sintaks Jinja2 menggunakan [def validate_template()](src/core/template_engine.py:93)
  - Pastikan semua variabel yang dipakai tersedia dari [def prepare_variables()](src/core/template_engine.py:33)

- Database/logs tidak muncul
  - Pastikan path database benar (panel kiri → Database File)
  - Pastikan aplikasi memiliki hak tulis di folder `database/` dan `logs/`

## 14. Keamanan & Praktik Baik

- Simpan kredensial SMTP secara aman (hindari commit ke VCS)
- Batasi akses write ke folder `config/`, `database/`, dan `logs/`
- Pertimbangkan enkripsi/secret manager untuk password produksi
- Gunakan whitelisting domain email untuk mencegah salah kirim

## 15. FAQ

- Apakah bisa menggunakan selain Outlook/Gmail?
  - Ya, selama mendukung SMTP dengan kredensial valid

- Apakah nama file harus mengikuti pola tertentu?
  - Ya, sesuai `key_pattern` yang mengandung identitas supplier (mis. `TT003`)

- Bisakah menambah variabel custom?
  - Ya, melalui panel Constant Variables (dua variabel custom yang dapat dipakai di subject/body)

- Bisakah melakukan scan awal file yang sudah ada?
  - Ya, tersedia API [def process_existing_files()](src/core/folder_monitor.py:104) (belum diekspos di UI)

## 16. Referensi Cepat Fungsi/Komponen

- Mulai aplikasi: [def main()](src/main.py:38)
- Start/Stop monitoring: [def start_monitoring()](src/ui/main_window.py:553), [def stop_monitoring()](src/ui/main_window.py:584)
- Kirim email uji: [def send_test_email()](src/ui/main_window.py:757)
- Preview email: [def generate_preview()](src/ui/main_window.py:907)
- Konfigurasi profile: [def get_profile_config()](src/core/config_manager.py:69), [def save_profile_config()](src/core/config_manager.py:98)
- Export/Import profile: [def export_profile()](src/core/config_manager.py:200), [def import_profile()](src/core/config_manager.py:206)
- Engine template: [def render_file_template()](src/core/template_engine.py:25), [def process_simple_variables()](src/core/template_engine.py:79)
- Monitoring: [class FolderMonitor](src/core/folder_monitor.py:47), [def move_file_to_sent()](src/core/folder_monitor.py:134)
- Pengirim email: [class EmailSenderFactory](src/core/email_sender.py:148), [class OutlookSender](src/core/email_sender.py:21), [class ThunderbirdSender](src/core/email_sender.py:73)

## 17. Contoh Pola Regex (Key Pattern)

- `([A-Z]{2}\d{3})` → cocok untuk `TT003_invoice_2024.pdf` (key = `TT003`)
- `DELIVERY_([A-Z0-9]+)` → cocok untuk `DELIVERY_AB12_2024.xlsx` (key = `AB12`)

Uji regex Anda dengan alat regex offline/online sebelum dipakai agar hasil ekstraksi sesuai.

## 18. Pengembangan & Pengujian

- Menjalankan tes dasar:
```bash
python test_app.py
```

- Menambahkan komponen baru:
  - Ikuti pola dan organisasi di folder `src/`
  - Tambahkan logging yang jelas
  - Perbarui dokumentasi profile bila menambah field baru di `config/default.ini`

## 19. Lisensi

Tambahkan lisensi yang sesuai di bagian “License”.

---