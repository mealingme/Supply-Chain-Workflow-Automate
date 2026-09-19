WAWI TRACKING UPLOAD - AUTOMATED MERGE
=======================================

WHAT THIS DOES
--------------
Reads today's UPS CSV and FedEx XLSX export from your Downloads folder and
merges them into one file: outbound_YYYYMMDD.csv, ready to upload into WAWI
(SOMS) to flip orders from NEW to SENT.

ONE-TIME SETUP
---------------
1. Install Python (if not already installed): https://www.python.org/downloads/
   During install, tick "Add python.exe to PATH".

2. Install the one extra library this script needs (reads the FedEx .xlsx file).
   Open Command Prompt and run:
       pip install openpyxl

3. Put these two files in the SAME folder, anywhere convenient (e.g. Desktop):
       merge_tracking.py
       run_tracking_merge.bat

DAILY USE
---------
1. Download today's UPS and FedEx files as normal into your Downloads folder
   (no need to rename anything).
2. Double-click run_tracking_merge.bat.
3. A black window opens, shows a short summary (how many rows found per
   courier, any warnings), and writes outbound_YYYYMMDD.csv into your
   Downloads folder.
4. Upload that file into WAWI as you do today.

WHAT IT CHECKS FOR YOU
-----------------------
- Skips FedEx rows that have an error or aren't OUTBOUND shipments (so you
  don't accidentally mark a failed shipment as SENT).
- Skips any row missing a reference or tracking number.
- Warns you if the same order reference shows up more than once (possible
  duplicate/mix-up between couriers).
- Warns you (rather than silently overwriting) if more than one UPS or
  FedEx file is sitting in Downloads for today - it uses the most recent one.
- Will NOT accidentally re-read a previous day's own output file as if it
  were a new UPS file, even though both filenames start with "outbound_".

DHL
---
Not wired up yet, since you have no DHL shipments today. When DHL resumes,
open merge_tracking.py and fill in the read_dhl() function near the bottom
(it already plugs straight into the merge - no other changes needed). Send
me a sample DHL export file and I can fill this in for you in under a
minute.

IF SOMETHING GOES WRONG
------------------------
- "No UPS/FedEx file found today": check the file is actually in Downloads
  and named the way it usually is (outbound_......csv / Shipment_Report_
  YYYY-MM-DD.xlsx).
- Duplicate reference warning: check WAWI manually for that order before
  uploading - don't upload blind if this fires.
- The window closes instantly / "python is not recognized": Python isn't
  installed or isn't on PATH - reinstall and tick "Add to PATH".
