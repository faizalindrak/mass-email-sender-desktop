import sqlite3
import json
from contextlib import contextmanager
import os

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

    def init_database(self):
        """Initialize database with required tables"""
        # Ensure database directory exists
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)

        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Create suppliers table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS suppliers (
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

            # Create email_logs table
            cursor.execute('''
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
                )
            ''')

            # Create configurations table
            cursor.execute('''
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
                    active BOOLEAN DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            conn.commit()

    def get_supplier_by_key(self, key):
        """Get supplier data by key"""
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

    def add_supplier(self, key, supplier_code, supplier_name, contact_name, emails, cc_emails=None, bcc_emails=None):
        """Add new supplier"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO suppliers
                (key, supplier_code, supplier_name, contact_name, emails, cc_emails, bcc_emails)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                key, supplier_code, supplier_name, contact_name,
                json.dumps(emails),
                json.dumps(cc_emails or []),
                json.dumps(bcc_emails or [])
            ))
            conn.commit()
            return cursor.lastrowid

    def log_email_sent(self, log_data):
        """Log sent email to database"""
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
            return cursor.lastrowid

    def get_email_logs(self, limit=100, offset=0):
        """Get email logs with pagination"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT * FROM email_logs
                ORDER BY sent_at DESC
                LIMIT ? OFFSET ?
            ''', (limit, offset))
            return [dict(row) for row in cursor.fetchall()]

    def get_all_suppliers(self, limit=None):
        """Get all active suppliers"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            query = "SELECT * FROM suppliers WHERE active = 1 ORDER BY supplier_name"
            if limit:
                query += f" LIMIT {limit}"
            cursor.execute(query)
            return [dict(row) for row in cursor.fetchall()]

    def update_supplier(self, supplier_id, **kwargs):
        """Update supplier data"""
        fields = []
        values = []

        for key, value in kwargs.items():
            if key in ['emails', 'cc_emails', 'bcc_emails'] and isinstance(value, list):
                value = json.dumps(value)
            fields.append(f"{key} = ?")
            values.append(value)

        if not fields:
            return

        values.append(supplier_id)

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE suppliers
                SET {', '.join(fields)}, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', values)
            conn.commit()

    def delete_supplier(self, supplier_id):
        """Soft delete supplier (set active = 0)"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE suppliers SET active = 0 WHERE id = ?", (supplier_id,))
            conn.commit()