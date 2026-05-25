@echo off
echo ============================================================
echo   POSTILION JOURNAL RECONCILIATION
echo   Step 2 of 2 — Matching PDFs against Excel ledger
echo ============================================================
echo.
py excel_check.py
echo.
echo ============================================================
echo   Done. Check ExcelFiles/ for highlighted reconciliation.
echo   Check logs/ for the reconciliation log.
echo ============================================================
cmd /k
