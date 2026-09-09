# Supply-Chain-Workflow-Automate

A small local Python tool that automates the daily cleanup of an
outstanding-shipments export before it's uploaded to an order
management system (referred to here as "SOMS" - the English
equivalent of the German term "WAWI"). It replaces a manual,
error-prone Excel editing routine with a deterministic script that runs
in seconds, entirely on your own machine.

No data ever leaves your computer — there's no network access, no
external API calls, no telemetry. It reads one local Excel file and
writes another.

## What it does

Each morning, a raw `.xlsx` export of outstanding customer shipments
needs a series of corrections applied before it's fit for import:
combining split order/receipt numbers, fixing postal codes, wrapping
overlong address lines to a carrier's character limit, correcting
prices that came out corrupted, and formatting phone number columns.

This script applies all of those rules automatically and adds a
**Notes and Warnings** column documenting exactly what it changed —
and, just as importantly, flags anything it *isn't* confident enough to
fix automatically, so a human reviews it instead of silently guessing.

Specifically, it:

- Combines `OrderNo` and `ReceiptNo` values (space-separated) for any
  order group that spans multiple line items, keeping the two columns
  separate.
- Pads short postal codes for a configurable set of countries with
  leading zeros.
- Removes duplicate address-line content that already appears in
  another field (city, postal code, country), then wraps address lines
  that exceed a configurable character limit across multiple fields at
  a safe word boundary — never overwriting existing data.
- Inserts a space between a letter and a directly-following digit in
  address fields (e.g. `FLAT16` → `FLAT 16`), without touching the
  reverse case (`2a`, a normal house-number suffix).
- Recalculates a corrupted price field as the net price before VAT,
  using a configurable table of country VAT rates, whenever the source
  value fails to parse as a plain number (for example, when it was
  accidentally read as a date).
- Writes designated phone-number columns as real numbers with a custom
  Excel display format, so they show a leading `+` without needing to
  be stored as text.
- Sorts the output by a configurable ID column.
- Stops and explains itself — rather than guessing — whenever it hits
  a case outside its configured rules (an unrecognized product, a
  missing rate/rule, more items in a group than it's configured to
  merge, more than one candidate input file, etc.).

## Requirements

- Python 3.10+
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install pandas openpyxl
```

## Usage

Place the source `.xlsx` file in the folder the script watches (by
default, your Downloads folder — see `DOWNLOADS_DIR` in the config
section), then run:

```bash
python OutstandingShipments.py
```

Or point it at a specific file:

```bash
python OutstandingShipments.py "/path/to/source.xlsx"
```

On Windows, `OutstandingShipments.bat` is a double-clickable launcher
that checks Python is installed and runs the script for you.

The script looks for a file matching the pattern configured in
`find_input_file()` (default: `OS_*.xlsx`), and writes its output
alongside it with `_for_soms` appended before the extension — e.g.
`OS_2026-08-27.xlsx` → `OS_2026-08-27_for_soms.xlsx`.

## Configuration

All business rules live in the `CONFIG` section at the top of
`OutstandingShipments.py` — column names, VAT rates, address-line
length limit, ZIP-padding countries, product-substitution mappings,
and so on. Adjusting behavior for a new rule, country, or edge case
should only require editing a value there, not the surrounding logic.

## Safety principles

- **Never guesses on ambiguous data.** Missing/placeholder postal
  codes, unrecognized products needing substitution, or prices in a
  country with no configured VAT rate all stop the script with a clear
  explanation rather than being silently invented.
- **Never processes the wrong file.** If zero or more than one
  candidate source file is found, the script stops rather than
  guessing which one is current.
- **Never overwrites data it shouldn't.** Address-field edits only
  happen when the target field is empty or genuinely duplicate content
  — anything else is flagged for manual review instead.

## License

No license has been specified for this repository — all rights
reserved by the author unless a license file is added.
