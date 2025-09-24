#!/usr/bin/env python3
"""
Email Automation Desktop - CSV Import Script

Usage:
    python database/import_db.py database/email_automation.db path/to/file.csv --table suppliers
    python database/import_db.py database/email_automation.db path/to/file.csv --table email_logs
    python database/import_db.py database/email_automation.db path/to/file.csv --table configurations

If --table is omitted, the script will attempt to auto-detect the target table based on CSV headers.
Supported tables: suppliers, email_logs, configurations

Notes:
- CSV must include headers.
- For array-like fields (emails, cc_emails, bcc_emails, recipient_emails), values may be JSON ('["a@x.com","b@y.com"]')
  or delimited text ("a@x.com; b@y.com" or "a@x.com,b@y.com"). They will be stored as JSON arrays in the database.
- For boolean fields (active, smtp_use_tls), accepted truthy values: 1, true, yes, y (case-insensitive).
- For suppliers and configurations, --update enables UPSERT on the unique keys (suppliers.key, configurations.profile_name).
"""

import argparse
import csv
import json
import os
import sqlite3
import sys
from typing import Dict, List, Tuple, Any, Optional


# -------------------------
# Table schemas (for import)
# -------------------------

TABLE_SCHEMAS: Dict[str, Dict[str, Any]] = {
    "suppliers": {
        "required": ["key", "supplier_code", "supplier_name", "emails"],
        "optional": ["contact_name", "cc_emails", "bcc_emails", "active"],
        "unique_key": "key",
        "array_fields": ["emails", "cc_emails", "bcc_emails"],
        "boolean_fields": ["active"],
        "transformers": {},
    },
    "email_logs": {
        "required": ["file_path", "filename", "supplier_key", "recipient_emails", "subject"],
        "optional": [
            "cc_emails",
            "bcc_emails",
            "body",
            "template_used",
            "status",
            "error_message",
            "email_client",
            "sent_at",
        ],
        "unique_key": None,
        "array_fields": ["recipient_emails", "cc_emails", "bcc_emails"],
        "boolean_fields": [],
        "transformers": {},
    },
    "configurations": {
        "required": ["profile_name", "monitor_folder", "sent_folder", "key_pattern", "email_client"],
        "optional": [
            "template_path",
            "subject_template",
            "body_template",
            "active",
            "created_at",
        ],
        "unique_key": "profile_name",
        "array_fields": [],
        "boolean_fields": ["active"],
        "transformers": {},
    },
}


# -------------------------
# Helpers
# -------------------------

def parse_bool(value: Any) -> Optional[int]:
    """Parse boolean-like values to integer 0/1 (SQLite friendly)."""
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in ("1", "true", "yes", "y", "on"):
        return 1
    if s in ("0", "false", "no", "n", "off", ""):
        return 0
    # If not recognizable, return None to let caller decide
    return None


def parse_array_cell(value: Any) -> Optional[str]:
    """Parse a cell that may represent an array (JSON or delimited) and return JSON string."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None

    # If already JSON array
    if s.startswith("[") and s.endswith("]"):
        try:
            arr = json.loads(s)
            if isinstance(arr, list):
                return json.dumps(arr, ensure_ascii=False)
        except Exception:
            # Fall through to delimiter parsing
            pass

    # Delimited by ; or ,
    delimiter = ";" if ";" in s else ("," if "," in s else None)
    if delimiter:
        parts = [p.strip() for p in s.split(delimiter) if p.strip()]
        return json.dumps(parts, ensure_ascii=False)

    # Single value -> wrap into array
    return json.dumps([s], ensure_ascii=False)


def auto_detect_table(headers: List[str]) -> Optional[str]:
    """Auto-detect table based on headers overlap with known schema."""
    header_set = {h.strip().lower() for h in headers if h}
    best_match = None
    best_score = -1

    for table, schema in TABLE_SCHEMAS.items():
        required = {c.lower() for c in schema["required"]}
        optional = {c.lower() for c in schema["optional"]}
        intersection = len(required & header_set)
        # Basic rule: must match all required to be valid
        if required.issubset(header_set):
            score = intersection + len(optional & header_set)
            if score > best_score:
                best_score = score
                best_match = table

    return best_match


def validate_headers(table: str, headers: List[str]) -> Tuple[bool, str]:
    """Validate that CSV headers contain at least required columns for the chosen table."""
    schema = TABLE_SCHEMAS[table]
    header_set = {h.strip() for h in headers if h}
    missing = [c for c in schema["required"] if c not in header_set]
    if missing:
        return False, f"Missing required columns for '{table}': {', '.join(missing)}"
    return True, "ok"


def transform_row(table: str, row: Dict[str, Any]) -> Dict[str, Any]:
    """Apply necessary type/format transformations per table schema."""
    schema = TABLE_SCHEMAS[table]
    result: Dict[str, Any] = {}

    # Only keep known columns (required + optional)
    allowed_cols = set(schema["required"]) | set(schema["optional"])
    for col in list(row.keys()):
        if col not in allowed_cols:
            # Ignore extra columns
            continue

        val = row[col]

        if col in schema["array_fields"]:
            result[col] = parse_array_cell(val)
        elif col in schema["boolean_fields"]:
            parsed = parse_bool(val)
            # If parse failed, try to coerce to int
            if parsed is None:
                try:
                    parsed = int(val)
                except Exception:
                    parsed = None
            result[col] = parsed
        else:
            # Keep string as-is
            result[col] = val if val is not None else None

    return result


def build_insert_sql(table: str, columns: List[str], update: bool) -> str:
    """Build INSERT (and optional UPSERT) SQL for a given table."""
    schema = TABLE_SCHEMAS[table]
    placeholders = ", ".join(["?"] * len(columns))
    cols_joined = ", ".join(columns)
    base_insert = f"INSERT INTO {table} ({cols_joined}) VALUES ({placeholders})"

    unique_key = schema.get("unique_key")
    if update and unique_key:
        # Construct UPSERT that updates all non-unique columns
        non_unique_cols = [c for c in columns if c != unique_key]
        set_clause_parts = []
        for c in non_unique_cols:
            if table == "suppliers" and c == "updated_at":
                # updated_at maintained via trigger; skip direct update
                continue
            set_clause_parts.append(f"{c}=excluded.{c}")
        set_clause = ", ".join(set_clause_parts) if set_clause_parts else ""
        if set_clause:
            upsert = f"{base_insert} ON CONFLICT({unique_key}) DO UPDATE SET {set_clause}"
        else:
            # If nothing to update, ignore conflict
            upsert = f"{base_insert} ON CONFLICT({unique_key}) DO NOTHING"
        return upsert

    return base_insert


def row_to_values(columns: List[str], row: Dict[str, Any]) -> List[Any]:
    """Map row dict into ordered values matching 'columns'."""
    values: List[Any] = []
    for c in columns:
        v = row.get(c, None)
        values.append(v)
    return values


# -------------------------
# Importer
# -------------------------

def import_csv(
    db_path: str,
    csv_path: str,
    table: Optional[str],
    delimiter: str = ",",
    encoding: str = "utf-8-sig",
    update: bool = False,
    batch_size: int = 500,
    dry_run: bool = False,
) -> Tuple[int, int, int]:
    """
    Import CSV into SQLite.

    Returns: (inserted_count, updated_count, error_count)
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    # Open CSV and read headers
    with open(csv_path, "r", encoding=encoding, newline="") as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        headers = reader.fieldnames or []
        if not headers:
            raise ValueError("CSV file must contain headers")

        # Decide table
        chosen_table = table or auto_detect_table(headers)
        if not chosen_table:
            raise ValueError(
                "Failed to auto-detect target table. "
                "Please specify --table {suppliers|email_logs|configurations}"
            )

        # Validate headers
        ok, msg = validate_headers(chosen_table, headers)
        if not ok:
            raise ValueError(msg)

        schema = TABLE_SCHEMAS[chosen_table]
        allowed_cols = set(schema["required"]) | set(schema["optional"])

        # Surviving columns for insert
        columns = [c for c in headers if c in allowed_cols]

        sql = build_insert_sql(chosen_table, columns, update)
        inserted = 0
        updated = 0
        errors = 0

        # Connect DB
        conn = sqlite3.connect(db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON;")
            cur = conn.cursor()

            batch_values: List[List[Any]] = []
            for raw_row in reader:
                try:
                    # Normalize None for empty strings
                    for k, v in list(raw_row.items()):
                        if isinstance(v, str):
                            v = v.strip()
                            raw_row[k] = v if v != "" else None

                    row = transform_row(chosen_table, raw_row)
                    values = row_to_values(columns, row)

                    if dry_run:
                        print(f"[DRY-RUN] {chosen_table}: {dict(zip(columns, values))}")
                        continue

                    batch_values.append(values)

                    if len(batch_values) >= batch_size:
                        cur.executemany(sql, batch_values)
                        conn.commit()
                        # Estimate inserts/updates
                        if update and schema.get("unique_key"):
                            updated += len(batch_values)  # conservative count (UPSERT may update or insert)
                        else:
                            inserted += len(batch_values)
                        batch_values.clear()

                except Exception as e:
                    errors += 1
                    print(f"[ERROR] Row import failed: {e}")

            # Flush remaining
            if batch_values and not dry_run:
                cur.executemany(sql, batch_values)
                conn.commit()
                if update and schema.get("unique_key"):
                    updated += len(batch_values)
                else:
                    inserted += len(batch_values)

        finally:
            conn.close()

        return inserted, updated, errors


# -------------------------
# CLI
# -------------------------

def main() -> int:
    parser = argparse.ArgumentParser(description="Import CSV into Email Automation SQLite DB.")
    parser.add_argument("db_file", help="Path to SQLite database file (e.g., database/email_automation.db)")
    parser.add_argument("csv_file", help="Path to CSV file to import")
    parser.add_argument(
        "--table",
        choices=list(TABLE_SCHEMAS.keys()),
        help="Target table name. If omitted, auto-detection is attempted.",
    )
    parser.add_argument("--delimiter", default=",", help="CSV delimiter (default: ',')")
    parser.add_argument("--encoding", default="utf-8-sig", help="CSV encoding (default: utf-8-sig)")
    parser.add_argument(
        "--update",
        action="store_true",
        help="Use UPSERT for tables with unique key (suppliers, configurations).",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=500,
        help="Number of rows per batch insert (default: 500).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print transformed rows without modifying the database.",
    )

    args = parser.parse_args()

    try:
        inserted, updated, errors = import_csv(
            db_path=args.db_file,
            csv_path=args.csv_file,
            table=args.table,
            delimiter=args.delimiter,
            encoding=args.encoding,
            update=args.update,
            batch_size=args.batch_size,
            dry_run=args.dry_run,
        )

        if args.dry_run:
            print("[DRY-RUN] Completed. No changes applied.")
        else:
            print(f"[DONE] Table: {args.table or 'auto-detected'} | Inserted: {inserted} | Upserted: {updated} | Errors: {errors}")

        return 0

    except Exception as e:
        print(f"[ERROR] Import failed: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())