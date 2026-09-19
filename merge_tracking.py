"""
merge_tracking.py

Merges daily UPS and FedEx tracking-number export files from the Downloads
folder into a single WAWI (SOMS)-ready CSV:

    outbound_YYYYMMDD.csv
    Package Reference No. 2;Tracking Number
    DN77726157;1Z54V67RDK46690339
    ...

Source files expected in the Downloads folder:
  - UPS:   outbound_<6 digits>_<long number>.csv
           e.g. outbound_091426_1789379413872.csv
           Uses columns: "Parcel Reference No. 2", "Tracking Number"
  - FedEx: Shipment_Report_YYYY-MM-DD.xlsx
           e.g. Shipment_Report_2026-09-14.xlsx
           Uses columns: "reference", "masterTrackingNumber"
           Rows are skipped if "errors" is non-blank or "shipmentType" is not
           OUTBOUND (failed / non-outbound shipments should not be marked SENT).
  - DHL:   not wired up yet. See read_dhl() below - add DHL logic there when
           DHL shipping resumes; it already plugs into the merge automatically.

Output:
  - outbound_YYYYMMDD.csv, saved to the same Downloads folder, semicolon
    delimited, matching the WAWI upload format exactly.

Run this by double-clicking run_tracking_merge.bat, or directly with:
    python merge_tracking.py
"""

import csv
import glob
import os
import re
import sys
from datetime import datetime

import openpyxl

# ----------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------

# Downloads folder (auto-detects the current Windows/Mac/Linux user's Downloads)
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

# Where the merged file gets written (same as Downloads, per your setup)
OUTPUT_DIR = DOWNLOADS_DIR

UPS_PATTERN = re.compile(r"^outbound_\d{6}_\d+\.csv$", re.IGNORECASE)
FEDEX_PATTERN = re.compile(r"^Shipment_Report_\d{4}-\d{2}-\d{2}\.xlsx$", re.IGNORECASE)
OUTPUT_PATTERN = re.compile(r"^outbound_\d{8}\.csv$", re.IGNORECASE)  # to avoid re-reading our own output


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------

def find_ups_files(folder):
    """Return UPS source CSVs in folder, newest first. Excludes our own output files."""
    matches = []
    for name in os.listdir(folder):
        if UPS_PATTERN.match(name) and not OUTPUT_PATTERN.match(name):
            matches.append(os.path.join(folder, name))
    matches.sort(key=os.path.getmtime, reverse=True)
    return matches


def find_fedex_files(folder, today_str):
    """Return FedEx source xlsx files. Prefer an exact match for today's date."""
    exact = os.path.join(folder, f"Shipment_Report_{today_str}.xlsx")
    if os.path.exists(exact):
        return [exact]
    matches = []
    for name in os.listdir(folder):
        if FEDEX_PATTERN.match(name):
            matches.append(os.path.join(folder, name))
    matches.sort(key=os.path.getmtime, reverse=True)
    return matches


def read_ups(path):
    """Read UPS CSV, return list of (reference, tracking_number) tuples, in file order."""
    rows = []
    skipped = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ref = (row.get("Parcel Reference No. 2") or "").strip()
            trk = (row.get("Tracking Number") or "").strip()
            if ref and trk:
                rows.append((ref, trk))
            else:
                skipped += 1
    return rows, skipped


def read_fedex(path):
    """Read FedEx xlsx, return list of (reference, tracking_number) tuples, in row order.
    Skips rows with a non-blank 'errors' value or shipmentType != OUTBOUND."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    headers = [str(h).strip() if h is not None else "" for h in header_row]
    col_idx = {h: i for i, h in enumerate(headers)}

    required = ["reference", "masterTrackingNumber"]
    missing = [c for c in required if c not in col_idx]
    if missing:
        raise ValueError(f"FedEx file is missing expected column(s): {missing}")

    has_errors_col = "errors" in col_idx
    has_type_col = "shipmentType" in col_idx

    rows = []
    skipped = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue

        ref = str(row[col_idx["reference"]] or "").strip()
        trk = str(row[col_idx["masterTrackingNumber"]] or "").strip()

        if has_errors_col:
            err_val = str(row[col_idx["errors"]] or "").strip()
            if err_val:
                skipped += 1
                continue

        if has_type_col:
            ship_type = str(row[col_idx["shipmentType"]] or "").strip().upper()
            if ship_type and ship_type != "OUTBOUND":
                skipped += 1
                continue

        if ref and trk:
            rows.append((ref, trk))
        else:
            skipped += 1

    return rows, skipped


def read_dhl():
    """Placeholder for DHL. Return list of (reference, tracking_number) tuples.
    Wire this up once DHL shipments resume - it will be included in the merge
    automatically, no other changes needed."""
    return [], 0


def write_output(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Package Reference No. 2", "Tracking Number"])
        for ref, trk in rows:
            writer.writerow([ref, trk])


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    today = datetime.now()
    today_iso = today.strftime("%Y-%m-%d")   # for FedEx filename match
    today_compact = today.strftime("%Y%m%d") # for output filename

    print(f"Looking in: {DOWNLOADS_DIR}")
    print(f"Today: {today_iso}\n")

    if not os.path.isdir(DOWNLOADS_DIR):
        print(f"ERROR: Downloads folder not found at {DOWNLOADS_DIR}")
        sys.exit(1)

    all_rows = []
    total_skipped = 0

    # --- UPS ---
    ups_files = find_ups_files(DOWNLOADS_DIR)
    if not ups_files:
        print("No UPS file found today (outbound_XXXXXX_XXXXXXXXXXXXXX.csv) - skipping UPS.")
    else:
        if len(ups_files) > 1:
            print(f"WARNING: found {len(ups_files)} UPS files, using the most recent:")
            for f in ups_files:
                print(f"   - {os.path.basename(f)}")
        ups_path = ups_files[0]
        ups_rows, ups_skipped = read_ups(ups_path)
        print(f"UPS   ({os.path.basename(ups_path)}): {len(ups_rows)} rows, {ups_skipped} skipped (missing ref/tracking)")
        all_rows.extend(ups_rows)
        total_skipped += ups_skipped

    # --- FedEx ---
    fedex_files = find_fedex_files(DOWNLOADS_DIR, today_iso)
    if not fedex_files:
        print("No FedEx file found today (Shipment_Report_YYYY-MM-DD.xlsx) - skipping FedEx.")
    else:
        if len(fedex_files) > 1:
            print(f"WARNING: found {len(fedex_files)} FedEx files, using the most recent:")
            for f in fedex_files:
                print(f"   - {os.path.basename(f)}")
        fedex_path = fedex_files[0]
        fedex_rows, fedex_skipped = read_fedex(fedex_path)
        print(f"FedEx ({os.path.basename(fedex_path)}): {len(fedex_rows)} rows, {fedex_skipped} skipped (errors/non-outbound/missing data)")
        all_rows.extend(fedex_rows)
        total_skipped += fedex_skipped

    # --- DHL (placeholder) ---
    dhl_rows, dhl_skipped = read_dhl()
    if dhl_rows:
        print(f"DHL: {len(dhl_rows)} rows, {dhl_skipped} skipped")
        all_rows.extend(dhl_rows)
        total_skipped += dhl_skipped

    if not all_rows:
        print("\nERROR: No tracking rows found from any courier. Nothing written.")
        sys.exit(1)

    # Duplicate reference check (would indicate a mix-up)
    refs = [r for r, _ in all_rows]
    dupes = {r for r in refs if refs.count(r) > 1}
    if dupes:
        print(f"\nWARNING: duplicate order reference(s) found across files: {sorted(dupes)}")
        print("Double-check these aren't the same order pulled from two couriers.")

    output_name = f"outbound_{today_compact}.csv"
    output_path = os.path.join(OUTPUT_DIR, output_name)

    if os.path.exists(output_path):
        print(f"\nWARNING: {output_name} already exists and will be overwritten.")

    write_output(output_path, all_rows)

    print(f"\nDone. Wrote {len(all_rows)} rows to:\n   {output_path}")
    if total_skipped:
        print(f"({total_skipped} source rows were skipped - see notes above)")


if __name__ == "__main__":
    main()
