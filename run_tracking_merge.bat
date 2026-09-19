@echo off
REM Double-click this file to build today's outbound_YYYYMMDD.csv
REM for WAWI (SOMS) upload from the UPS and FedEx files in your Downloads folder.

cd /d "%~dp0"
python merge_tracking.py

echo.
pause
