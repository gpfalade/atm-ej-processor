@echo off
echo ============================================================
echo   POSTILION JOURNAL PROCESSOR
echo   Step 1 of 2 — Processing .dat files and generating PDFs
echo ============================================================
echo.
py journal_split.py
echo.
echo ============================================================
echo   Done. Check output_folder/ for generated PDFs.
echo   Check logs/ for the processing log.
echo   Run run_excel_check.bat to reconcile against Excel.
echo ============================================================
cmd /k
