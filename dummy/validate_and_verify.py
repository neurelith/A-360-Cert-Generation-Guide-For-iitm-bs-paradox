#!/usr/bin/env python3
"""
Certificate Validation & Verification Utility for Saavan '26
-----------------------------------------------------------
Performs pre-flight dataset integrity audits, UUID validation,
verification URL verification, Firestore payload compilation,
and generated PDF reconciliation.

Zero secrets in source code. Do NOT commit .env or service credentials.
"""

import os
import sys
import re
import json
import argparse
import pandas as pd
from typing import Dict, List, Tuple

# Try to import config if available
try:
    import config
    DEFAULT_CSV = config.DUMMY_DATA_CSV
    VALID_TYPES = list(config.PPTX_TEMPLATES.keys())
except ImportError:
    DEFAULT_CSV = os.path.join(os.path.dirname(__file__), "dummy_data.csv")
    VALID_TYPES = ["winner", "participant", "judge", "guest"]

UUID_V4_REGEX = re.compile(
    r"^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$",
    re.IGNORECASE,
)
EMAIL_REGEX = re.compile(
    r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
)

REQUIRED_COLUMNS = [
    "cert_id",
    "type",
    "department",
    "event_name",
    "name",
    "email",
    "position",
    "uuid",
    "qr_code",
]


def audit_csv(csv_path: str) -> Tuple[bool, pd.DataFrame, List[str]]:
    """Runs strict sanity and schema audits on the input CSV."""
    errors = []
    warnings = []

    if not os.path.exists(csv_path):
        return False, pd.DataFrame(), [f"CSV file not found: {csv_path}"]

    try:
        df = pd.read_csv(csv_path, dtype=str)
    except Exception as e:
        return False, pd.DataFrame(), [f"Failed to parse CSV: {e}"]

    # Clean whitespace in column headers
    df.columns = [c.strip() for c in df.columns]

    # 1. Column presence check
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing mandatory columns: {missing_cols}")
        return False, df, errors

    total_rows = len(df)
    if total_rows == 0:
        errors.append("Dataset is empty (0 rows).")
        return False, df, errors

    print(f"[*] Auditing {total_rows} rows from {os.path.basename(csv_path)}...")

    # 2. Check for missing values in mandatory columns
    for col in REQUIRED_COLUMNS:
        null_indices = df[df[col].isna() | (df[col].astype(str).str.strip() == "")].index.tolist()
        if null_indices:
            errors.append(f"Column '{col}' has empty values at row index(es): {[i + 2 for i in null_indices]}")

    # 3. Uniqueness checks
    for col in ["cert_id", "uuid", "qr_code"]:
        duplicates = df[df.duplicated(subset=[col], keep=False)]
        if not duplicates.empty:
            dup_vals = duplicates[col].unique().tolist()
            errors.append(f"Duplicate values found in '{col}': {dup_vals}")

    # 4. UUID & QR code format checks
    for idx, row in df.iterrows():
        row_num = idx + 2
        raw_uuid = str(row.get("uuid", "")).strip()
        qr_url = str(row.get("qr_code", "")).strip()
        email = str(row.get("email", "")).strip()
        cert_type = str(row.get("type", "")).strip().lower()

        # Validate UUID
        if not UUID_V4_REGEX.match(raw_uuid):
            errors.append(f"Row {row_num}: Invalid UUID format '{raw_uuid}' (Must match RFC 4122 uppercase hex)")

        # Validate QR contains the matching UUID
        if raw_uuid.upper() not in qr_url.upper():
            errors.append(f"Row {row_num}: QR URL '{qr_url}' does not contain candidate UUID '{raw_uuid}'")

        # Validate Email
        if not EMAIL_REGEX.match(email):
            warnings.append(f"Row {row_num}: Suspect email format '{email}'")

        # Validate Type routing
        if cert_type not in VALID_TYPES:
            errors.append(f"Row {row_num}: Unknown cert type '{cert_type}'. Expected one of: {VALID_TYPES}")

    if warnings:
        print("\n[!] Warnings:")
        for w in warnings:
            print(f"    - {w}")

    is_valid = len(errors) == 0
    return is_valid, df, errors


def generate_firestore_payload(df: pd.DataFrame, season: str = "Saavan '26", output_path: str = None) -> Dict:
    """
    Transforms the candidate records into the exact Firestore collection ('certs') schema
    used by the certificate verification portal (https://verify.iitmparadox.org).
    
    Document ID: UUID without hyphens or uppercase UUID
    Fields:
      - name: str
      - date: str (e.g. 'May 2026')
      - roll: str (email / roll number)
      - role: str (position / category)
      - season: str
    """
    records = {}
    for _, row in df.iterrows():
        raw_uuid = str(row["uuid"]).strip().replace("/", "").upper()
        records[raw_uuid] = {
            "name": str(row["name"]).strip(),
            "date": "May 2026",
            "roll": str(row["email"]).strip(),
            "role": str(row["position"]).strip(),
            "season": season,
            "cert_id": str(row["cert_id"]).strip(),
            "event": str(row["event_name"]).strip(),
        }

    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=2)
        print(f"[+] Exported {len(records)} verification records to: {output_path}")

    return records


def reconcile_generated_pdfs(df: pd.DataFrame, pdf_dir: str) -> Tuple[int, List[str]]:
    """Checks whether compiled vector PDFs exist for every candidate row."""
    missing = []
    matched = 0

    if not os.path.exists(pdf_dir):
        return 0, [f"PDF output directory not found: {pdf_dir}"]

    existing_files = set(os.listdir(pdf_dir))

    for _, row in df.iterrows():
        cid = str(row["cert_id"]).strip()
        ctype = str(row["type"]).strip().lower()
        expected_filename = f"saavan26_{ctype}_{cid}.pdf"
        if expected_filename in existing_files:
            matched += 1
        else:
            missing.append(expected_filename)

    return matched, missing


def main():
    parser = argparse.ArgumentParser(description="Saavan '26 Certificate Verification & Validation Tool")
    parser.add_argument("--csv", default=DEFAULT_CSV, help="Path to input candidate CSV")
    parser.add_argument("--export-firestore", metavar="OUTPUT_JSON", help="Export Firestore certs batch JSON")
    parser.add_argument("--check-pdfs", action="store_true", help="Reconcile generated PDFs against CSV")
    parser.add_argument("--pdf-dir", default=None, help="Directory containing generated PDFs")
    args = parser.parse_args()

    print("=" * 65)
    print("  SAAVAN '26 — CERTIFICATE VALIDATION & VERIFICATION RUNNER")
    print("=" * 65)

    is_valid, df, errors = audit_csv(args.csv)

    if not is_valid:
        print("\n[x] Dataset validation FAILED with the following errors:")
        for err in errors:
            print(f"    - {err}")
        sys.exit(1)

    print("\n[v] Dataset Integrity: 100% PASS (All schemas, UUIDs & URLs verified)")

    if args.export_firestore:
        generate_firestore_payload(df, output_path=args.export_firestore)

    if args.check_pdfs:
        pdf_dir = args.pdf_dir
        if not pdf_dir:
            pdf_dir = os.path.join(os.path.dirname(args.csv), "generated_certificates", "pdf")
        matched, missing = reconcile_generated_pdfs(df, pdf_dir)
        print(f"\n[*] PDF Reconciliation against {pdf_dir}:")
        print(f"    - Matched: {matched}/{len(df)} PDFs")
        if missing:
            print(f"    [!] Missing {len(missing)} PDFs:")
            for m in missing[:10]:
                print(f"        - {m}")
            if len(missing) > 10:
                print(f"        ... and {len(missing) - 10} more")
        else:
            print("    [v] 100% of PDFs accounted for.")

    print("\n[+] Verification & validation check complete.\n")


if __name__ == "__main__":
    main()
