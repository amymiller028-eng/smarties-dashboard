@echo off
REM Double-click this file to refresh data.json from the Excel workbook.
cd /d "%~dp0"
echo Reading Program-Evals-2026.xlsx and rebuilding data.json...
echo.
python compute-data.py
echo.
echo Done. Now upload data.json to GitHub.
pause
