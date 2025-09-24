#!/usr/bin/env python3
"""
Email Automation Desktop - Database Initialization Script

Usage:
    python database/init_db.py database/schema.sql database/email_automation.db
    python database/init_db.py path/to/sqlcommand.sql path/to/mydb.db
    python database/init_db.py database/schema.sql database/email_automation.db --overwrite

This script initializes an SQLite database using a provided SQL schema file.
"""

import argparse
import sqlite3
import os
import sys


def read_sql_file(path: str) -> str:
    """Read SQL file content"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"SQL file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def ensure_parent_dir(path: str) -> None:
    """Ensure parent directory for db file exists"""
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)


def init_db(sql_path: str, db_path: str) -> None:
    """Initialize SQLite DB from schema SQL"""
    sql_text = read_sql_file(sql_path)
    ensure_parent_dir(db_path)

    # Connect and execute schema
    conn = sqlite3.connect(db_path)
    try:
        # Recommended pragmas
        conn.execute("PRAGMA foreign_keys = ON;")
        # Execute the full schema script
        conn.executescript(sql_text)
        conn.commit()
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Initialize SQLite database from SQL schema."
    )
    parser.add_argument("sql_file", help="Path to SQL schema file")
    parser.add_argument("db_file", help="Path to SQLite database file to create/use")
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Delete existing DB file before init (use with caution)",
    )
    args = parser.parse_args()

    sql_file = args.sql_file
    db_file = args.db_file

    try:
        # Optionally remove existing DB for a clean init
        if args.overwrite and os.path.exists(db_file):
            os.remove(db_file)
            print(f"[INFO] Removed existing database: {db_file}")

        init_db(sql_file, db_file)

        print(f"[SUCCESS] Database initialized: {os.path.abspath(db_file)}")
        print(f"[INFO] Applied schema from: {os.path.abspath(sql_file)}")
        return 0

    except Exception as e:
        print(f"[ERROR] Initialization failed: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())