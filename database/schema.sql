-- Email Automation Desktop - SQLite Schema
PRAGMA foreign_keys = ON;

-- Suppliers table
CREATE TABLE IF NOT EXISTS suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT UNIQUE NOT NULL,
    supplier_code TEXT NOT NULL,
    supplier_name TEXT NOT NULL,
    contact_name TEXT,
    emails TEXT NOT NULL,
    cc_emails TEXT,
    bcc_emails TEXT,
    active INTEGER DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_suppliers_key ON suppliers(key);

-- Trigger to auto-update updated_at on suppliers
CREATE TRIGGER IF NOT EXISTS trg_suppliers_updated_at
AFTER UPDATE ON suppliers
FOR EACH ROW
BEGIN
    UPDATE suppliers SET updated_at = CURRENT_TIMESTAMP WHERE id = OLD.id;
END;

-- Email logs table
CREATE TABLE IF NOT EXISTS email_logs (
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
);

CREATE INDEX IF NOT EXISTS idx_email_logs_supplier_key ON email_logs(supplier_key);
CREATE INDEX IF NOT EXISTS idx_email_logs_sent_at ON email_logs(sent_at);

-- Configurations table
CREATE TABLE IF NOT EXISTS configurations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    profile_name TEXT UNIQUE NOT NULL,
    monitor_folder TEXT NOT NULL,
    sent_folder TEXT NOT NULL,
    key_pattern TEXT NOT NULL,
    email_client TEXT NOT NULL,
    template_path TEXT,
    subject_template TEXT,
    body_template TEXT,
    active INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Seed example (optional). Comment out if not needed.
-- INSERT INTO suppliers (key, supplier_code, supplier_name, contact_name, emails)
-- VALUES ('TT003', 'TT003', 'TOKO TOKO ABADI', 'Budi Santoso', '["budi@tokoabadi.com"]');