# Shipping Label Batch-Upload Generator

A small Python workflow that takes a daily shipment export from a Sales
Order Management System (**SOMS**) and an order/product-detail export, and
turns them into the three courier batch-upload files needed to print
shipping labels:

- **UK Fedex** upload file (recipient country `GB`) - includes full UK
  customs data (HS/TARIC codes, country of origin, weight, English
  description, declared value), since the UK is outside the EU.
- **EU Fedex** upload file (everything else, e.g. `IE`, `ES`, `IT`...).
- **UPS** upload file (recipient country `FR` or `DE`).

No customs data is added to the EU/UPS files, since intra-EU shipments
don't need a customs declaration.

## Why this exists

Manually splitting a day's orders by destination country, reformatting
them into each courier's required spreadsheet layout, and - for UK orders
- manually looking up HS codes, country of origin, and per-product
customs values is slow and error-prone. This script automates all of that
from files you already have.

## How it works, at a glance

1. Read the day's SOMS shipment export (recipient/shipping details, one row
   per shipment).
2. Route each shipment by recipient country: `FR`/`DE` → UPS, `GB` → UK
   Fedex, everything else → EU Fedex.
3. For UK shipments, look up each order's products from the order/product
   export, and each product's customs data (HS code, country of origin,
   weight, English description) from the product master. Multiple distinct
   products in one order become FedEx's documented multi-commodity rows;
   the same product appearing more than once in an order is summed into a
   single line (real repeat-quantity purchases, not export duplicates -
   see the comments in `load_products_by_order()` if you're curious why
   that distinction matters).
4. Write the three output files, using each template's own layout/styling,
   and highlight (yellow) any cell where required data couldn't be found
   rather than guessing.

## Folder layout

| Location | Contents |
|---|---|
| **Your Downloads folder** (or wherever `SHIP_DOWNLOAD_DIR` points) | Only the two files that change every day: the SOMS shipment export and the order/product export. The finished output files are also written here. |
| **This script's own folder** (or wherever `SHIP_TEMPLATE_DIR` points - defaults to next to `create_shipping_uploads.py`) | Everything that does **not** change daily: the UK/EU Fedex templates, the UPS template, the product master, and `product_description_translations.json`. |

Expected filenames (case/spacing-tolerant - see [Naming quirks](#naming-quirks-the-script-tolerates) below):

```
Downloads/
  WAWI_Sendungsdaten_<date>.xlsx      # today's shipment export
  OS_<date>_for_wawi.xlsx             # today's order/product export

<script folder>/
  create_shipping_uploads.py
  run_shipping_uploads.bat
  UK Fedex Shipments Upload Template.xlsx
  EU Fedex Shipments Upload Template.xlsx
  UPS_Sendungsdaten_<date>.xlsx       # the UPS template
  Products_<timestamp>.xlsx           # product master
  product_description_translations.json
```

> **Note on filenames:** the shipment/order export filenames above still
> say `WAWI` / `wawi` because that's literally what the source system
> names them - the code labels this system "SOMS" (Sales Order Management
> System) internally, but the glob patterns that find your real files were
> deliberately left matching the real filenames. If your source system's
> export naming ever changes, update `SOMS_GLOB` and `OS_GLOB` in
> `create_shipping_uploads.py` (each has a comment marking it).

## Setup

1. **Install Python 3** if you don't have it (`run_shipping_uploads.bat`
   will tell you if it's missing, with a link).
2. Put `create_shipping_uploads.py`, `run_shipping_uploads.bat`,
   `product_description_translations.json`, the two Fedex templates, the
   UPS template, and the product master all in the same folder.
3. Open `run_shipping_uploads.bat` in a text editor and check
   `SHIP_DOWNLOAD_DIR` points at your real Downloads folder.
4. Double-click `run_shipping_uploads.bat`. First run installs `openpyxl`
   automatically if needed.

### Do the templates need real data in them, or are headers enough?

- **UK/EU Fedex templates: need at least one real data row (row 2).** The
  script copies that row's columns `AE:AO` (account number, terms of sale,
  notification settings, etc.) onto every output row. Headers alone means
  those columns come out blank.
- **UPS template: headers alone are fine.** Every column there is
  shipment-specific and gets fully overwritten; a data row only affects
  cosmetic formatting (font, etc.) of the output.

## Customizing output filenames

Output files are named after their template with the word `Template`
replaced by a configurable label (default: `Processed`), e.g.
`UK Fedex Shipments Upload Template.xlsx` → `UK Fedex Shipments Upload Processed.xlsx`.
The UPS output instead gets `_Processed` appended (same configurable
label). To use a different label (e.g. to match what your team already
expects), set the `SHIP_OUTPUT_LABEL` environment variable - there's a
commented-out example line in `run_shipping_uploads.bat` showing how.
**Don't commit a real person's name there if this repo is public or
shared** - keep that line local/uncommitted.

## Naming quirks the script tolerates

Real-world exports are messy. The script is deliberately forgiving about:

- **Spaces vs. underscores** in the UK/EU Fedex template filenames
  (`UK Fedex Shipments Upload Template.xlsx` and
  `UK_Fedex_Shipments_Upload_Template.xlsx` both work).
- **Stray whitespace** around order numbers when matching the shipment
  export to the order/product export (a trailing space on one side used
  to silently break the match).
- **Its own previous output files** - a glob search for the UPS/SOMS
  template will never accidentally pick up a `_Processed` (or whatever
  your label is) file this script wrote on an earlier run, even if that
  file is now the most recently modified match.

When something genuinely can't be found, the script prints the **exact,
quoted contents** of the relevant folder so you can spot things like a
hidden double extension (`file.xlsx.xlsx`) or a trailing space at a
glance, instead of guessing from a stack trace.

## Data quality assumptions (please sanity-check these)

- `customsValue` = the **net (ex-VAT)** unit price from the order export,
  summed across repeat lines for the same product in one order. If your
  courier/customs broker expects VAT-inclusive value instead, this needs
  changing.
- English customs descriptions come from
  `product_description_translations.json`, which maps the product master's
  own German descriptions (material/composition/function - the same facts
  a customs declaration needs) to English, **not** from the public
  webshop's marketing copy (which uses promotional/wellness language
  that's inappropriate for a customs form). Add new products' descriptions
  there as they come up; missing ones fall back to the product name and
  get highlighted yellow.
- `manufacturingCountry` comes from the product master's "Country of
  Origin" column, converted to an ISO-2 code via `COUNTRY_NAME_TO_CODE` in
  the script. A country name that isn't in that dict yet is left blank +
  highlighted rather than guessed - add it there if that happens.
- `commodityWeight` is the product master's per-unit weight multiplied by
  the (summed) quantity for that line.

If a product's SKU isn't in the product master at all, or the master file
is missing entirely, the relevant customs cells are left blank and
highlighted yellow rather than guessed.

## Privacy / what's (deliberately) not in this repo

This repo processes real customer shipment data, but the **script itself**
doesn't hardcode any of it - customer names, addresses, emails, and phone
numbers only ever exist transiently in the Excel files you feed it at
runtime, never in the code.

A couple of things were still cleaned up before publishing:

- **A colleague's real first name** was previously hardcoded as the output
  filename marker. It's now a configurable `SHIP_OUTPUT_LABEL` environment
  variable (default `"Processed"`) instead - see
  [Customizing output filenames](#customizing-output-filenames).
- **The company's brand name**, which appeared in one product's English
  customs description (in `product_description_translations.json`), has
  been replaced with a generic `[Brand]` placeholder in this published
  copy. The German source key was left untouched so the lookup still works
  against your real product master - only the English output text was
  redacted. **If you rely on that specific product**, restore the real
  brand name in your local copy before using it for actual customs
  paperwork, or that one item's description will print `[Brand]` on the
  label.

If you're publishing your own fork, it's worth double-checking
`product_description_translations.json` and the product master for
anything else specific to your business before making the repo public.

## Requirements

- Python 3.8+
- `openpyxl` (installed automatically by `run_shipping_uploads.bat` if
  missing; otherwise `pip install openpyxl`)

## Files in this repo

| File | Purpose |
|---|---|
| `create_shipping_uploads.py` | The workflow itself. |
| `run_shipping_uploads.bat` | Windows double-click launcher. |
| `product_description_translations.json` | German → English customs description lookup, editable without touching code. |
