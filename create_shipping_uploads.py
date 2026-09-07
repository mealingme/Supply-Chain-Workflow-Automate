#!/usr/bin/env python3
"""
create_shipping_uploads.py
===========================

Daily workflow that takes the SOMS shipment export ("WAWI_Sendungsdaten_*.xlsx")
and splits/reformats it into the three courier batch-upload files:

    - UK_Fedex_Shipments_Upload{SUFFIX}.xlsx   (recipient country == GB)
    - EU_Fedex_Shipments_Upload{SUFFIX}.xlsx   (everything else, e.g. IE, ES, IT ...)
    - UPS_Sendungsdaten_<date>{SUFFIX}.xlsx    (recipient country FR or DE)

Business rules implemented (per instructions):
    1. UK orders (Land == 'GB')           -> UK Fedex file
    2. FR and DE orders (Land in FR, DE)  -> UPS file
    3. All other orders                   -> EU Fedex file
    4. UPS file: SOMS columns A:AC are copied across unchanged (same layout).
    5. Fedex files (UK & EU):
         - SOMS columns A:AC map letter-for-letter onto the same columns in
           the Fedex file (col C -> col C, col H -> col H, etc.)
         - Column A is always left BLANK in the Fedex output files.
         - Column AD (recipient notification language code) = the recipient's
           country code, EXCEPT:
               GB -> en
               IE -> en
               BE -> fr
               AT -> DE   (see NOTE below - user said "AU"; see note)
         - Columns AE:AO are constant on every row. The constant values are
           read from the existing template file itself (row 2) so that if
           those business constants (account numbers, terms of sale, etc.)
           are ever updated in the template, this script will pick them up
           automatically rather than relying on hardcoded values.
    6. UK Fedex file ONLY - customs / product data (GB is not in the EU, so
       every UK parcel needs a customs declaration; intra-EU parcels don't):
         - Product line items are matched onto each shipment from the
           "OS_*_for_wawi.xlsx" order-detail export via the DN tracking
           number (OS column "OrderNo" == SOMS "Referenz2").
         - The same product code can appear on several rows of an order -
           these are genuinely separate units (each independently priced,
           e.g. 12.71/12.71/12.71/12.70 - not identical copies of one
           row), so they are SUMMED into a single commodity line per
           product code: Quantity = sum of each row's quantity,
           customsValue = sum of each row's net price (i.e. the total
           value for that quantity, which is what FedEx's customsValue
           field expects).
         - Each product's customs facts (HS/TARIC code, country of
           manufacture, per-unit weight, description) come from the
           product master file ("Products_*.xlsx"), matched on SKU == OS
           ProdCode; commodityWeight is that per-unit weight multiplied by
           the line's summed quantity. This master is the authoritative
           source - nothing here is guessed.
         - The master's "Product description" column is German marketing/
           internal text. product_description_translations.json (next to
           this script) maps each exact German description to a short
           English, customs-relevant description (material/composition/
           function only - no marketing language). Add new entries there
           as new products appear; no code changes needed. If a product's
           German description has no translation on file yet, the product
           name is used as a fallback and the cell is highlighted yellow
           for manual review.
         - For a shipment with ONE product, its details are written
           straight onto the shipment's own row (itemDescription,
           commodityQuantity, commodityMeasureUnit, commodityWeight,
           customsValue, manufacturingCountry, harmonizedCode).
         - For a shipment with MULTIPLE distinct products, the first
           product's details go on the main row and one extra "COMMODITY"
           row (recordType = COMMODITY) is added per additional product,
           all sharing the same shipmentReference (the DN number) - this
           is FedEx's documented multi-commodity format.
         - customsValue uses the net (ex-VAT) unit price from the OS file,
           currencyType = EUR, purposeOfShipment = SOLD, generateInvoice = CI.
         - If a product's SKU isn't found in the product master at all (or
           the master file is missing), the relevant cells are left blank
           and highlighted yellow rather than guessed.

    ASSUMPTIONS TO CONFIRM:
        - customsValue = net product price (ex-VAT) from the OS file. If
          FedEx / HMRC expect the VAT-inclusive price instead, change
          CUSTOMS_VALUE_FIELD below.
        - English descriptions were deliberately built by translating the
          product master's own German descriptions (factual: material,
          composition, product type) rather than pulling copy from the
          public webshop, which is written in promotional/wellness
          language (e.g. energy/biofield claims) that isn't appropriate
          for a customs declaration. Please sanity-check a few entries in
          product_description_translations.json against how your customs
          broker likes descriptions phrased.
        - manufacturingCountry is taken from the product master's "Country
          of Origin" column and converted to an ISO-2 code via
          COUNTRY_NAME_TO_CODE below. If a country name shows up that
          isn't in that dict yet, it's left blank + highlighted rather
          than guessed - add it to the dict when that happens.

    NOTE on "AU": confirmed with the user - "AU" in the original
    instructions meant Austria (country code AT), not Australia. The
    script maps AT -> DE accordingly.

Output files are saved into DOWNLOAD_DIR with the same base filename as the
existing template files, with SUFFIX (e.g. "_Processed", configurable via
OUTPUT_LABEL below) appended before the .xlsx extension, e.g.
"UK_Fedex_Shipments_Upload_Processed.xlsx".

FOLDER LAYOUT:
    - DOWNLOAD_DIR (defaults to your Downloads folder): only the two daily
      downloads go here - WAWI_Sendungsdaten_<date>.xlsx and
      OS_<date>_for_wawi.xlsx. This is also where the finished
      files are written.
    - TEMPLATE_DIR (defaults to this script's own folder, i.e. wherever
      you keep create_shipping_uploads.py and run_shipping_uploads.bat):
      everything that doesn't change daily lives here instead -
      UK/EU Fedex templates, the UPS_Sendungsdaten_<date>.xlsx template,
      the Products_<timestamp>.xlsx product master, and
      product_description_translations.json. Keep all of these next to
      the script and .bat file; nothing needs to go in Downloads except
      the two files above.

HOW TO ADAPT FOR PRODUCTION:
    - Point DOWNLOAD_DIR at the real Windows/Mac Downloads folder.
    - Point TEMPLATE_DIR wherever you keep the courier upload-format files
      and product master, if not next to the script.
    - Once you're happy, wire OUTPUT_DIR to the courier apps' watched
      "batch upload" folders instead of DOWNLOAD_DIR, and this script can
      run unattended.
"""

import glob
import json
import os
import re
import sys
from collections import OrderedDict
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill

# --------------------------------------------------------------------------
# CONFIG - edit these for your environment
# --------------------------------------------------------------------------

# The folder this script itself lives in. Keep the courier template files
# (UK/EU Fedex, UPS Sendungsdaten), the product master, and
# product_description_translations.json all here, next to the .py and .bat
# files - they're reference/business data, not daily downloads, so they
# don't belong in the Downloads folder.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Folder the SOMS export and OS order/product export are downloaded into
# each day, and where the finished output files should be written.
DOWNLOAD_DIR = os.environ.get("SHIP_DOWNLOAD_DIR", "/mnt/user-data/outputs")

# Folder holding the courier upload-format template files, the product
# master, and product_description_translations.json. Defaults to this
# script's own folder - override with the SHIP_TEMPLATE_DIR environment
# variable if you'd rather keep them somewhere else.
TEMPLATE_DIR = os.environ.get("SHIP_TEMPLATE_DIR", SCRIPT_DIR)

# Word used to mark this script's output files, e.g. output files are named
# "UK Fedex Shipments Upload <OUTPUT_LABEL>.xlsx". The original workflow
# used a colleague's first name here to mark files as
# machine-processed and ready for that person to sort - that's a real
# person's name, so it isn't hardcoded in this public script. Set it via
# the SHIP_OUTPUT_LABEL environment variable (e.g. in run_shipping_uploads.bat,
# NOT committed to a public repo) if you want your real output filenames to
# match what your team already expects; otherwise it defaults to "Processed".
OUTPUT_LABEL = os.environ.get("SHIP_OUTPUT_LABEL", "Processed")
SUFFIX = f"_{OUTPUT_LABEL}"

UK_TEMPLATE_NAME = "UK Fedex Shipments Upload Template.xlsx"
EU_TEMPLATE_NAME = "EU Fedex Shipments Upload Template.xlsx"
UPS_TEMPLATE_GLOB = "UPS_Sendungsdaten_*.xlsx"
SOMS_GLOB = "WAWI_Sendungsdaten_*.xlsx"
# ^ NOTE: the variable is named SOMS_GLOB, but its VALUE is left as
# "WAWI_Sendungsdaten_*.xlsx" deliberately - that's the literal filename
# your source system actually produces each day. If you rename that export
# (or switch systems), update this one line to match the new filename -
# don't rename it just to be consistent with the rest of the code, or the
# script will stop finding your daily file.

# Order/product detail export used to add customs data to the UK Fedex file.
OS_GLOB = "OS_*_for_wawi.xlsx"
# ^ same note as SOMS_GLOB above: "_for_wawi" here is part of the real
# filename your order system produces, not a label - leave it as-is unless
# that export's naming actually changes.

# --- Customs constants for the UK Fedex file (see ASSUMPTIONS in the module
#     docstring - confirm these with the business before relying on them) ---
CUSTOMS_CURRENCY = "EUR"
CUSTOMS_PURPOSE_OF_SHIPMENT = "SOLD"
CUSTOMS_GENERATE_INVOICE = "CI"
CUSTOMS_COMMODITY_TYPE = "ITEMS"
CUSTOMS_MEASURE_UNIT = "PCS"

# Product master export (SKU -> TARIC/HS code, country of origin, weight,
# description) used to fill in UK customs data.
PRODUCT_MASTER_GLOB = "Products_*.xlsx"

# Country-of-origin names (as they appear in the product master) -> ISO-2
# code. Add to this if a new country of origin shows up in the master.
COUNTRY_NAME_TO_CODE = {
    "Australia": "AU",
    "Canada": "CA",
    "China": "CN",
    "France": "FR",
    "Germany": "DE",
    "Hong Kong": "HK",
    "Italy": "IT",
    "Kazakhstan": "KZ",
    "Malaysia": "MY",
    "South Korea": "KR",
    "Switzerland": "CH",
    "Thailand": "TH",
    "United States": "US",
}

# German -> English customs description lookup, kept in a separate JSON
# file next to this script so it can be extended without touching code.
DESCRIPTION_TRANSLATIONS_PATH = os.path.join(SCRIPT_DIR, "product_description_translations.json")

MISSING_DATA_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Country -> recipient notification language code exceptions.
# Default (not listed here) = language code equals the country code itself.
LANGUAGE_EXCEPTIONS = {
    "GB": "en",
    "IE": "en",
    "BE": "fr",
    "AT": "DE",  # see NOTE in module docstring re: "AU" -> interpreted as Austria
}

# Countries routed to the UPS file (all others that aren't GB go to EU Fedex)
UPS_COUNTRIES = {"FR", "DE"}
UK_COUNTRY = "GB"

FEDEX_LAST_COL = 29  # column AC (A=1 ... AC=29): SOMS data columns copied across
FEDEX_CONST_FIRST_COL = 31  # column AE
FEDEX_CONST_LAST_COL = 41  # column AO

# Customs / product columns (UK Fedex file only)
COL_ITEM_DESCRIPTION = 27  # AA
COL_GENERATE_INVOICE = 42  # AP
COL_PURPOSE_OF_SHIPMENT = 43  # AQ
COL_CURRENCY_TYPE = 44  # AR
COL_SHIPMENT_REFERENCE = 45  # AS
COL_COMMODITY_TYPE = 46  # AT
COL_RECORD_TYPE = 47  # AU
COL_HARMONIZED_CODE = 48  # AV
COL_MANUFACTURING_COUNTRY = 49  # AW
COL_COMMODITY_QUANTITY = 50  # AX
COL_COMMODITY_MEASURE_UNIT = 51  # AY
COL_COMMODITY_WEIGHT = 52  # AZ
COL_CUSTOMS_VALUE = 53  # BA
UK_LAST_COL = 53


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def find_one(directory, pattern, label):
    matches = _glob_excluding_own_output(directory, pattern)
    if not matches:
        raise FileNotFoundError(
            f"Could not find {label} matching '{pattern}' in {directory}"
        )
    # Most-recently-modified match wins (handles dated filenames)
    return max(matches, key=os.path.getmtime)


def find_first_match(directory, pattern):
    """Like find_one, but returns None instead of raising - used for preflight checks."""
    matches = _glob_excluding_own_output(directory, pattern)
    return max(matches, key=os.path.getmtime) if matches else None


def _glob_excluding_own_output(directory, pattern):
    """
    glob.glob(), but excluding this script's own previous output files
    outputs. Without this, a glob like 'UPS_Sendungsdaten_*.xlsx' would
    match both the real template/source file AND an output file this
    script wrote on an earlier run - and since that output file is newer,
    "most recently modified wins" would pick our own output as the input,
    which is wrong.
    """
    matches = glob.glob(os.path.join(directory, pattern))
    return [m for m in matches if f"{SUFFIX}." not in os.path.basename(m)]


def extract_date_suffix(filename):
    """Pull a trailing _YYYYMMDD (or similar) date token out of a filename, if present."""
    m = re.search(r"(\d{8})", os.path.basename(filename))
    return m.group(1) if m else None


def load_soms_rows(path):
    """Read SOMS data rows (columns A:AC) as a list of lists, skipping blank rows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, FEDEX_LAST_COL + 1)]
        if all(v is None or v == "" for v in values):
            continue
        rows.append(values)
    return rows


def classify(row):
    """Return 'UPS', 'UK', or 'EU' based on the Land (country) column (index 7 = col H)."""
    country_code = (row[7] or "").strip().upper()
    if country_code in UPS_COUNTRIES:
        return "UPS", country_code
    if country_code == UK_COUNTRY:
        return "UK", country_code
    return "EU", country_code


def language_code_for(country_code):
    return LANGUAGE_EXCEPTIONS.get(country_code, country_code)


def load_products_by_order(path):
    """
    Read the OS_*_for_wawi.xlsx order/product-detail export and return
    {order_no: [ {ProdCode, ProdName, Quantity, Price}, ... ]}

    Each distinct product code within an order becomes ONE line, but when
    the same product code appears on several rows of an order, those rows
    are genuinely separate units (each independently priced/rounded - e.g.
    12.71, 12.71, 12.71, 12.70 - not identical copies of one row), so they
    are SUMMED rather than discarded:
      - Quantity = sum of each row's Quantity (normally 1 each)
      - Price    = sum of each row's ProductPrice, i.e. the TOTAL net
                    value for that quantity of that product (this is what
                    FedEx's customsValue field expects - "total customs
                    value of the commodity", not a per-unit price).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}

    required = ["OrderNo", "ProdCode", "ProdName", "Quantity", "ProductPrice"]
    missing = [h for h in required if h not in headers]
    if missing:
        raise ValueError(f"OS file is missing expected column(s): {missing}")

    def to_number(value, cast):
        try:
            return cast(str(value).strip())
        except (TypeError, ValueError):
            return cast(0)

    products_by_order = {}
    for r in range(2, ws.max_row + 1):
        order_no = ws.cell(row=r, column=headers["OrderNo"]).value
        if not order_no:
            continue
        order_no = str(order_no).strip()  # source data sometimes has stray trailing spaces
        prod_code = ws.cell(row=r, column=headers["ProdCode"]).value
        prod_name = ws.cell(row=r, column=headers["ProdName"]).value
        quantity = to_number(ws.cell(row=r, column=headers["Quantity"]).value, int)
        price = to_number(ws.cell(row=r, column=headers["ProductPrice"]).value, float)

        bucket = products_by_order.setdefault(order_no, OrderedDict())
        if prod_code in bucket:
            bucket[prod_code]["Quantity"] += quantity
            bucket[prod_code]["Price"] += price
        else:
            bucket[prod_code] = {
                "ProdCode": prod_code,
                "ProdName": prod_name,
                "Quantity": quantity,
                "Price": round(price, 2),
            }

    for order_no, bucket in products_by_order.items():
        for entry in bucket.values():
            entry["Price"] = round(entry["Price"], 2)

    # Convert OrderedDicts to plain lists, preserving first-seen order.
    return {order_no: list(items.values()) for order_no, items in products_by_order.items()}


def load_description_translations():
    try:
        with open(DESCRIPTION_TRANSLATIONS_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"WARNING: description translation file not found: {DESCRIPTION_TRANSLATIONS_PATH}")
        return {}
    data.pop("_readme", None)
    return data


def load_product_master(path, description_translations):
    """
    Read the product master (Products_*.xlsx) and return
    {sku: {name, taric, country_code, country_raw, weight, description, description_is_fallback}}

    Skips rows missing both a TARIC code and a country of origin (test/junk
    entries and internal-only items with no export data).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    required = ["Product name", "SKU", "TARIC Code", "Country of Origin", "Weight", "Product description"]
    missing = [h for h in required if h not in headers]
    if missing:
        raise ValueError(f"Product master file is missing expected column(s): {missing}")

    master = {}
    for r in range(2, ws.max_row + 1):
        sku = ws.cell(row=r, column=headers["SKU"]).value
        if sku is None:
            continue
        taric = ws.cell(row=r, column=headers["TARIC Code"]).value
        country_raw = ws.cell(row=r, column=headers["Country of Origin"]).value
        if not taric or not country_raw:
            continue  # test/junk row or item with no export data

        name = ws.cell(row=r, column=headers["Product name"]).value
        weight = ws.cell(row=r, column=headers["Weight"]).value
        desc_de = ws.cell(row=r, column=headers["Product description"]).value

        description = description_translations.get(desc_de) if desc_de else None
        description_is_fallback = description is None
        if description is None:
            description = name  # fall back to the (English) product name

        master[str(sku)] = {
            "name": name,
            "taric": re.sub(r"[^0-9]", "", str(taric)),
            "country_code": COUNTRY_NAME_TO_CODE.get(country_raw),
            "country_raw": country_raw,
            "weight": weight,
            "description": description,
            "description_is_fallback": description_is_fallback,
        }
    return master


def clear_data_rows(ws, header_row=1, max_col=53):
    """Wipe out any existing sample/data rows below the header, values only."""
    if ws.max_row <= header_row:
        return
    for r in range(header_row + 1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).value = None


def style_from(ws, row, col):
    cell = ws.cell(row=row, column=col)
    return {
        "font": copy(cell.font),
        "number_format": cell.number_format,
        "alignment": copy(cell.alignment),
        "border": copy(cell.border),
        "fill": copy(cell.fill),
    }


def apply_style(cell, style):
    cell.font = style["font"]
    cell.number_format = style["number_format"]
    cell.alignment = style["alignment"]
    cell.border = style["border"]
    cell.fill = style["fill"]


# --------------------------------------------------------------------------
# Builders
# --------------------------------------------------------------------------

def build_ups_file(template_path, ups_rows, out_path):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    # Capture reference style from the template's first data row (row 2)
    # before we wipe it, so new rows keep the same look & feel.
    ref_row = 2 if ws.max_row >= 2 else None
    col_styles = None
    if ref_row:
        col_styles = {c: style_from(ws, ref_row, c) for c in range(1, FEDEX_LAST_COL + 1)}

    clear_data_rows(ws, header_row=1, max_col=FEDEX_LAST_COL)

    for i, row in enumerate(ups_rows, start=2):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=i, column=c, value=value)
            if col_styles:
                apply_style(cell, col_styles[c])

    wb.save(out_path)
    return len(ups_rows)


def build_fedex_file(template_path, rows_with_country, out_path):
    """rows_with_country: list of (soms_row_values, country_code)"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    # Capture the AE:AO "constant across all rows" values from the template
    # itself (row 2) before wiping, per the "these stay the same" instruction.
    const_values = {}
    if ws.max_row >= 2:
        for c in range(FEDEX_CONST_FIRST_COL, FEDEX_CONST_LAST_COL + 1):
            const_values[c] = ws.cell(row=2, column=c).value

    # Capture reference style (font etc.) per column from row 2 as well.
    col_styles = None
    if ws.max_row >= 2:
        col_styles = {
            c: style_from(ws, 2, c) for c in range(1, FEDEX_CONST_LAST_COL + 1)
        }

    clear_data_rows(ws, header_row=1, max_col=ws.max_column)

    for i, (row, country_code) in enumerate(rows_with_country, start=2):
        # Column A left blank deliberately.
        # Columns B:AC = SOMS columns B:AC, copied letter-for-letter.
        for c in range(2, FEDEX_LAST_COL + 1):
            value = row[c - 1]  # row is 0-indexed, col A=1 -> index 0
            cell = ws.cell(row=i, column=c, value=value)
            if col_styles:
                apply_style(cell, col_styles[c])

        # Column AD: recipient notification language code
        ad_cell = ws.cell(row=i, column=FEDEX_CONST_FIRST_COL - 1, value=language_code_for(country_code))
        if col_styles:
            apply_style(ad_cell, col_styles[FEDEX_CONST_FIRST_COL - 1])

        # Columns AE:AO: constant values taken from the template
        for c in range(FEDEX_CONST_FIRST_COL, FEDEX_CONST_LAST_COL + 1):
            cell = ws.cell(row=i, column=c, value=const_values.get(c))
            if col_styles:
                apply_style(cell, col_styles[c])

    wb.save(out_path)
    return len(rows_with_country)


def build_uk_fedex_file(template_path, rows_with_country, products_by_order, product_master, out_path):
    """
    Like build_fedex_file, but also writes product/customs data (needed
    because GB is outside the EU). rows_with_country: list of
    (soms_row_values, country_code). products_by_order comes from
    load_products_by_order(), keyed by the DN order number (SOMS col W).
    product_master comes from load_product_master(), keyed by SKU/ProdCode.
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    const_values = {}
    if ws.max_row >= 2:
        for c in range(FEDEX_CONST_FIRST_COL, FEDEX_CONST_LAST_COL + 1):
            const_values[c] = ws.cell(row=2, column=c).value

    col_styles = None
    if ws.max_row >= 2:
        col_styles = {c: style_from(ws, 2, c) for c in range(1, UK_LAST_COL + 1)}

    clear_data_rows(ws, header_row=1, max_col=max(ws.max_column, UK_LAST_COL))

    stats = {
        "shipments": 0,
        "products_matched": 0,
        "no_order_data": 0,
        "no_master_data": 0,
        "missing_country_code": 0,
        "fallback_description": 0,
    }

    def write_cell(row_idx, col_idx, value, highlight=False):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        if col_styles and col_idx in col_styles:
            apply_style(cell, col_styles[col_idx])
        if highlight:
            cell.fill = MISSING_DATA_FILL
        return cell

    r = 2
    for row, country_code in rows_with_country:
        stats["shipments"] += 1
        dn = row[22]  # column W (Referenz2) = DN order number, 0-indexed -> 22
        if dn is not None:
            dn = str(dn).strip()
        products = products_by_order.get(dn, [])
        if not products:
            stats["no_order_data"] += 1

        # --- main shipment row: all the usual shipment fields ---
        for c in range(2, FEDEX_LAST_COL + 1):
            write_cell(r, c, row[c - 1])
        write_cell(r, FEDEX_CONST_FIRST_COL - 1, language_code_for(country_code))  # AD
        for c in range(FEDEX_CONST_FIRST_COL, FEDEX_CONST_LAST_COL + 1):
            write_cell(r, c, const_values.get(c))

        # --- customs / product fields ---
        write_cell(r, COL_SHIPMENT_REFERENCE, dn)
        if products:
            write_cell(r, COL_GENERATE_INVOICE, CUSTOMS_GENERATE_INVOICE)
            write_cell(r, COL_PURPOSE_OF_SHIPMENT, CUSTOMS_PURPOSE_OF_SHIPMENT)
            write_cell(r, COL_CURRENCY_TYPE, CUSTOMS_CURRENCY)

            def write_commodity(row_idx, product, is_main_row):
                stats["products_matched"] += 1
                master_entry = product_master.get(str(product["ProdCode"]))
                if not master_entry:
                    stats["no_master_data"] += 1

                description = (master_entry or {}).get("description") or product["ProdName"]
                if not master_entry or master_entry.get("description_is_fallback"):
                    stats["fallback_description"] += 1
                country_code_out = (master_entry or {}).get("country_code")
                if master_entry and not country_code_out:
                    stats["missing_country_code"] += 1
                taric = (master_entry or {}).get("taric")
                unit_weight = (master_entry or {}).get("weight")
                total_weight = (
                    round(unit_weight * product["Quantity"], 3)
                    if unit_weight is not None and product["Quantity"]
                    else None
                )

                write_cell(row_idx, COL_ITEM_DESCRIPTION, description, highlight=not master_entry or master_entry.get("description_is_fallback"))
                write_cell(row_idx, COL_COMMODITY_TYPE, CUSTOMS_COMMODITY_TYPE)
                write_cell(row_idx, COL_HARMONIZED_CODE, taric, highlight=not taric)
                write_cell(row_idx, COL_MANUFACTURING_COUNTRY, country_code_out, highlight=not country_code_out)
                write_cell(row_idx, COL_COMMODITY_QUANTITY, product["Quantity"])
                write_cell(row_idx, COL_COMMODITY_MEASURE_UNIT, CUSTOMS_MEASURE_UNIT)
                write_cell(row_idx, COL_COMMODITY_WEIGHT, total_weight, highlight=total_weight is None)
                write_cell(row_idx, COL_CUSTOMS_VALUE, product["Price"])
                if not is_main_row:
                    write_cell(row_idx, COL_SHIPMENT_REFERENCE, dn)
                    write_cell(row_idx, COL_RECORD_TYPE, "COMMODITY")

            # First product goes directly on the shipment row.
            write_commodity(r, products[0], is_main_row=True)
            r += 1

            # Extra products (if any) get their own COMMODITY rows.
            for product in products[1:]:
                write_commodity(r, product, is_main_row=False)
                r += 1
        else:
            # No product data found for this DN - flag it so it's obvious a
            # manual customs entry is needed before this shipment can go out.
            write_cell(r, COL_ITEM_DESCRIPTION, "NEEDS PRODUCT DATA", highlight=True)
            r += 1

    wb.save(out_path)
    return stats


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def find_named_file(directory, expected_name):
    """
    Look for a file called exactly expected_name in directory. If that
    exact name isn't there, also accept a version that only differs by
    spaces vs underscores (e.g. "UK Fedex Shipments Upload.xlsx" instead
    of "UK_Fedex_Shipments_Upload.xlsx") - both are reasonable ways to
    name the same template file. Returns the real path found, or None.
    """
    exact = os.path.join(directory, expected_name)
    if os.path.isfile(exact):
        return exact
    target_norm = expected_name.replace(" ", "_").lower()
    try:
        entries = os.listdir(directory)
    except OSError:
        return None
    for entry in entries:
        if entry.replace(" ", "_").lower() == target_norm:
            return os.path.join(directory, entry)
    return None


def list_directory(directory):
    """
    Full directory listing with each filename wrapped in repr() so
    invisible issues (trailing spaces, hidden double extensions like
    '.xlsx.xlsx', smart-quotes, etc.) are impossible to miss.
    """
    try:
        entries = sorted(os.listdir(directory))
    except OSError as exc:
        return [f"  (could not list this folder: {exc})"]
    if not entries:
        return ["  (folder is empty)"]
    return [f"  {entry!r}" for entry in entries]


def main():
    # --- Preflight: check every REQUIRED file exists before doing any work.
    # (OS order/product export and the product master are OPTIONAL - handled
    # separately below with warnings, since the script can still produce a
    # UK Fedex file without them, just without customs product data.)
    problems = []

    soms_path = find_first_match(DOWNLOAD_DIR, SOMS_GLOB)
    if not soms_path:
        problems.append(
            f"  - SOMS export: no file matching '{SOMS_GLOB}' in\n"
            f"      {DOWNLOAD_DIR}\n"
            f"    (this is today's shipment export - download it from SOMS first)"
        )

    uk_template = find_named_file(TEMPLATE_DIR, UK_TEMPLATE_NAME)
    if not uk_template:
        problems.append(
            f"  - UK Fedex template file not found:\n      {os.path.join(TEMPLATE_DIR, UK_TEMPLATE_NAME)}\n"
            f"    (this is a reference/template file, not a daily download - "
            f"save a copy of '{UK_TEMPLATE_NAME}' into that folder)"
        )

    eu_template = find_named_file(TEMPLATE_DIR, EU_TEMPLATE_NAME)
    if not eu_template:
        problems.append(
            f"  - EU Fedex template file not found:\n      {os.path.join(TEMPLATE_DIR, EU_TEMPLATE_NAME)}\n"
            f"    (this is a reference/template file, not a daily download - "
            f"save a copy of '{EU_TEMPLATE_NAME}' into that folder)"
        )

    ups_template = find_first_match(TEMPLATE_DIR, UPS_TEMPLATE_GLOB)
    if not ups_template:
        problems.append(
            f"  - UPS template: no file matching '{UPS_TEMPLATE_GLOB}' in\n"
            f"      {TEMPLATE_DIR}\n"
            f"    (this is a reference/template file, not a daily download)"
        )

    if problems:
        print("ERROR: required file(s) are missing - cannot continue.\n")
        print("\n\n".join(problems))
        print(
            "\nNote: the UK/EU Fedex templates and the UPS template only need to be "
            "saved once (they don't change daily) - they define the column layout "
            "the courier apps expect. Re-run this script once they're all in place."
        )
        print(
            f"\nHere's exactly what this script can see in {DOWNLOAD_DIR}\n"
            f"(each name in quotes - check closely for things like a hidden double\n"
            f"extension '.xlsx.xlsx', a trailing space, or spaces instead of underscores):\n"
        )
        for line in list_directory(DOWNLOAD_DIR):
            print(line)
        if TEMPLATE_DIR != DOWNLOAD_DIR:
            print(f"\nAnd in {TEMPLATE_DIR}:\n")
            for line in list_directory(TEMPLATE_DIR):
                print(line)
        return 1

    print(f"SOMS source file : {soms_path}")
    print(f"UK Fedex template: {uk_template}")
    print(f"EU Fedex template: {eu_template}")
    print(f"UPS template     : {ups_template}")

    # Order/product detail file for UK customs data. Optional-ish: if it's
    # missing, the UK file is still produced but every shipment will be
    # flagged as NEEDS PRODUCT DATA.
    try:
        os_path = find_one(DOWNLOAD_DIR, OS_GLOB, "OS order/product export file")
        print(f"OS product file  : {os_path}")
        products_by_order = load_products_by_order(os_path)
    except FileNotFoundError:
        print(f"WARNING: no OS order/product export found matching '{OS_GLOB}' in {DOWNLOAD_DIR}.")
        print("         The UK Fedex file will be built WITHOUT customs product data.")
        products_by_order = {}

    # Product master (SKU -> HS/TARIC code, country of origin, weight, description).
    description_translations = load_description_translations()
    try:
        product_master_path = find_one(TEMPLATE_DIR, PRODUCT_MASTER_GLOB, "product master file")
        print(f"Product master   : {product_master_path}")
        product_master = load_product_master(product_master_path, description_translations)
        print(f"  -> {len(product_master)} products with usable customs data (TARIC + country of origin).")
    except FileNotFoundError:
        print(f"WARNING: no product master found matching '{PRODUCT_MASTER_GLOB}' in {TEMPLATE_DIR}.")
        print("         UK customs fields (HS code, country of origin, weight) will be left blank + highlighted.")
        product_master = {}

    soms_rows = load_soms_rows(soms_path)
    print(f"\nRead {len(soms_rows)} shipment rows from SOMS file.")

    ups_rows, uk_rows, eu_rows = [], [], []
    for row in soms_rows:
        dest, country_code = classify(row)
        if dest == "UPS":
            ups_rows.append(row)
        elif dest == "UK":
            uk_rows.append((row, country_code))
        else:
            eu_rows.append((row, country_code))

    print(f"  -> UPS (FR/DE): {len(ups_rows)} rows")
    print(f"  -> UK Fedex (GB): {len(uk_rows)} rows")
    print(f"  -> EU Fedex (other): {len(eu_rows)} rows")

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    def out_name(template_path):
        """Same base name as template, + SUFFIX (used for the UPS file)."""
        base = os.path.basename(template_path)
        name, ext = os.path.splitext(base)
        return os.path.join(DOWNLOAD_DIR, f"{name}{SUFFIX}{ext}")

    def out_name_replace_template(template_path):
        """
        UK/EU Fedex files: same name as the template, but with the word
        'Template' replaced by OUTPUT_LABEL (e.g. 'UK Fedex Shipments
        Upload Template.xlsx' -> 'UK Fedex Shipments Upload Processed.xlsx').
        Falls back to appending SUFFIX if the actual filename doesn't
        contain 'Template' (e.g. an older-style template name).
        """
        base = os.path.basename(template_path)
        name, ext = os.path.splitext(base)
        if "Template" in name:
            new_name = name.replace("Template", OUTPUT_LABEL)
        else:
            new_name = f"{name}{SUFFIX}"
        return os.path.join(DOWNLOAD_DIR, f"{new_name}{ext}")

    uk_out = out_name_replace_template(uk_template)
    eu_out = out_name_replace_template(eu_template)
    ups_out = out_name(ups_template)

    n_ups = build_ups_file(ups_template, ups_rows, ups_out)
    uk_stats = build_uk_fedex_file(uk_template, uk_rows, products_by_order, product_master, uk_out)
    n_eu = build_fedex_file(eu_template, eu_rows, eu_out)

    print("\nWritten:")
    print(f"  {ups_out}  ({n_ups} rows)")
    print(
        f"  {uk_out}  ({uk_stats['shipments']} shipments, "
        f"{uk_stats['products_matched']} product lines)"
    )
    print(f"  {eu_out}  ({n_eu} rows)")

    if uk_stats["no_order_data"]:
        print(
            f"\nWARNING: {uk_stats['no_order_data']} UK shipment(s) had no matching "
            f"order in the OS file - marked 'NEEDS PRODUCT DATA' in the output, "
            f"highlighted yellow."
        )
    if uk_stats["no_master_data"]:
        print(
            f"WARNING: {uk_stats['no_master_data']} product line(s) have no matching SKU "
            f"in the product master - HS code/country/weight left blank + highlighted."
        )
    if uk_stats["missing_country_code"]:
        print(
            f"WARNING: {uk_stats['missing_country_code']} product line(s) have a country "
            f"of origin in the product master that isn't in COUNTRY_NAME_TO_CODE yet - "
            f"add it there. Highlighted yellow in the output."
        )
    if uk_stats["fallback_description"]:
        print(
            f"WARNING: {uk_stats['fallback_description']} product line(s) have no English "
            f"customs description on file - used the product name as a fallback, "
            f"highlighted yellow. Add a translation to "
            f"product_description_translations.json to fix."
        )


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001 - deliberately broad: this is the top-level CLI entry point
        print(f"\nERROR: something went wrong and the run stopped early.\n  {type(exc).__name__}: {exc}")
        print("\nIf this keeps happening, share this message so it can be fixed.")
        sys.exit(1)
