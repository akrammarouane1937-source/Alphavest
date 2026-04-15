@echo off
cd /d "c:\Users\àf\Downloads\School project"

:: Ensure logs folder exists
if not exist logs mkdir logs

echo [%date% %time%] Starting daily update >> logs\daily.log 2>&1

:: Step 1 — fetch new closing prices and compute MASI
python update_prices_bvc.py >> logs\update_prices.log 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: update_prices_bvc.py failed >> logs\daily.log 2>&1
    exit /b 1
)

:: Step 2 — recompute all risk metrics and portfolio metrics
python daily_automation.py >> logs\daily_automation.log 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: daily_automation.py failed >> logs\daily.log 2>&1
    exit /b 1
)

:: Step 3 — copy updated Excel files to data/ (for Streamlit app)
if not exist data mkdir data
copy /y "risk_metrics_by_window_latest.xlsx"  "data\" >> logs\daily.log 2>&1
copy /y "portfolio_metrics_latest.xlsx"        "data\" >> logs\daily.log 2>&1

:: Step 4 — copy updated Excel files to OneDrive (shared with professor)
set ONEDRIVE="C:\Users\àf\OneDrive\Alphavest - Rapport PFE"
copy /y "risk_metrics_by_window_latest.xlsx"  %ONEDRIVE% >> logs\daily.log 2>&1
copy /y "portfolio_metrics_latest.xlsx"        %ONEDRIVE% >> logs\daily.log 2>&1

echo [%date% %time%] Daily update completed — files synced to data/ and OneDrive >> logs\daily.log 2>&1
