"""
Price Update Script — casablanca-bourse.com API
=================================================
Downloads missing closing prices from the Casablanca Stock Exchange
public API (api.casablanca-bourse.com) and appends them to
risk_metrics_full_with_masi.csv.

How it works:
  - Closing price = last transaction price (executedPrice) of each trading day
  - MASI index values are read from the local CSV (Moroccan All Shares - Données Historiques.csv)
  - 67/77 stock tickers have confirmed data; remaining few may have no transactions

Usage:
    python update_prices_bvc.py

No API key or login required.
"""

import pandas as pd
import numpy as np
import requests
import time
import os
import sys
import warnings

warnings.filterwarnings('ignore')

# ============================================================
# CONFIG
# ============================================================
BASE_DIR      = r'c:\Users\àf\Downloads\School project'
PRICES_FILE   = os.path.join(BASE_DIR, 'historical_prices.csv')
BACKUP_FILE   = os.path.join(BASE_DIR, 'historical_prices_BACKUP.csv')
WEIGHTS_FILE  = os.path.join(BASE_DIR, 'Compo_All_Indices_20260408.xls')
API_BASE      = 'https://api.casablanca-bourse.com'
MASI_COL      = 'MASI'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json',
    'Referer': 'https://www.casablanca-bourse.com/',
}

# ============================================================
# STEP 1 — Load existing price file
# ============================================================
print("Loading existing price file...")
prices = pd.read_csv(PRICES_FILE, sep=';', dayfirst=True,
                     parse_dates=['Date'], index_col='Date', dtype=object)
prices.index = pd.to_datetime(prices.index, dayfirst=True, errors='coerce')
for col in prices.columns:
    prices[col] = prices[col].astype(str).str.replace(',', '.', regex=False).str.strip()
prices = prices.apply(pd.to_numeric, errors='coerce').sort_index()

last_date = prices.index.max()
print(f"  Current file ends: {last_date.date()}")
print(f"  Tickers: {len(prices.columns)} (including MASI)")

# ============================================================
# STEP 2 — Determine date range
# ============================================================
start_dt = last_date + pd.Timedelta(days=1)
end_dt   = pd.Timestamp.today()

print(f"\nTarget range: {start_dt.date()} to {end_dt.date()}")

# Stock tickers in the file (strip .CS suffix for API queries)
stock_tickers = [c for c in prices.columns if c != MASI_COL]
ticker_map = {t.replace('.CS', ''): t for t in stock_tickers}  # ATW -> ATW.CS

# ============================================================
# STEP 3 — Get trading dates in range
# ============================================================
def get_trading_dates(start_dt, end_dt):
    """Get list of dates that had any transactions (= trading days)."""
    dates = pd.date_range(start_dt, end_dt, freq='B')  # business days estimate
    # Filter to actual trading days by checking if any transactions exist
    trading = []
    for d in dates:
        date_str = d.strftime('%Y-%m-%d')
        try:
            r = requests.get(
                f'{API_BASE}/fr/api/bourse_data/transaction'
                f'?filter[transactTime][value]={date_str}'
                f'&filter[transactTime][operator]=CONTAINS'
                f'&page[limit]=1',
                headers=HEADERS, timeout=8, verify=False)
            data = r.json()
            count = data.get('meta', {}).get('count', 0) or 0
            if count > 0:
                trading.append(d)
                print(f"  {date_str}: {count} transactions — TRADING DAY")
            else:
                print(f"  {date_str}: 0 transactions — holiday/weekend")
        except Exception as e:
            print(f"  {date_str}: error — {str(e)[:50]}")
        time.sleep(0.2)
    return trading

print("\nIdentifying trading days...")
trading_dates = get_trading_dates(start_dt, end_dt)
print(f"\nFound {len(trading_dates)} trading days.")

if not trading_dates:
    print("No new trading days found. File is up to date.")
    sys.exit(0)

# ============================================================
# STEP 4 — Download closing prices for each ticker/date
# ============================================================
def get_closing_price(symbol, date_str):
    """Closing price = last executedPrice of the day (sort by transactTime desc)."""
    try:
        r = requests.get(
            f'{API_BASE}/fr/api/bourse_data/transaction'
            f'?filter[transactTime][value]={date_str}'
            f'&filter[transactTime][operator]=CONTAINS'
            f'&filter[symbol.symbol]={symbol}'
            f'&sort=-transactTime'
            f'&page[limit]=1',
            headers=HEADERS, timeout=10, verify=False)
        data = r.json()
        items = data.get('data', [])
        if items:
            price = items[0]['attributes']['executedPrice']
            return float(price) if price is not None else None
        return None
    except Exception:
        return None

print(f"\nDownloading closing prices for {len(stock_tickers)} tickers x {len(trading_dates)} days...")
print(f"Estimated time: ~{len(stock_tickers) * len(trading_dates) * 0.25 / 60:.1f} minutes\n")

new_data = {}  # {ticker_with_cs: {date: price}}

for i, (api_sym, file_col) in enumerate(ticker_map.items()):
    prices_for_ticker = {}
    any_data = False

    for d in trading_dates:
        date_str = d.strftime('%Y-%m-%d')
        price = get_closing_price(api_sym, date_str)
        if price is not None:
            prices_for_ticker[d] = price
            any_data = True
        time.sleep(0.15)

    if any_data:
        new_data[file_col] = prices_for_ticker
        sample = list(prices_for_ticker.items())[:2]
        print(f"  [{i+1:2d}/{len(ticker_map)}] {file_col:<10} — {len(prices_for_ticker)} prices | e.g. {sample}")
    else:
        print(f"  [{i+1:2d}/{len(ticker_map)}] {file_col:<10} — no data (illiquid or delisted)")

# ============================================================
# STEP 5 — MASI calculated from weighted stock returns
# ============================================================
print("\nCalculating MASI from weighted stock returns...")

# ISIN -> BVC symbol mapping (pre-fetched; stable unless BVC changes ISINs)
ISIN_TICKER = {
    'MA0000012783': 'GTM',  # SGTM S.A
    'MA0000012585': 'AKT',  # AKDITAL
    'MA0000012627': 'CFG',  # CFG BANK
    'MA0000012718': 'CMG',  # CMGP GROUP
    'MA0000012759': 'VCN',  # VICENNE
    'MA0000012767': 'CAP',  # CASH PLUS S.A
}

# Load MASI composition weights
weights_df = pd.read_excel(WEIGHTS_FILE)
masi_w = weights_df[weights_df['Indice'] == 'MASI'][['Code ISIN', 'Poids']].copy()
masi_w.columns = ['isin', 'weight']

# For ISINs not in our hardcoded dict, look up dynamically from BVC API
for isin in masi_w['isin']:
    if isin not in ISIN_TICKER:
        try:
            r = requests.get(
                f'{API_BASE}/fr/api/bourse_data/instrument?filter[codeISIN]={isin}&page[limit]=1',
                headers=HEADERS, timeout=5, verify=False)
            data = r.json()
            if data.get('data'):
                ISIN_TICKER[isin] = data['data'][0]['attributes']['symbol']
        except Exception:
            pass
        time.sleep(0.05)

masi_w['ticker'] = masi_w['isin'].map(ISIN_TICKER)

# Fetch prices for GTM, VCN, CAP (in MASI but not in main prices CSV)
extra_masi = {'GTM': {}, 'VCN': {}, 'CAP': {}}
for sym in extra_masi:
    for d in trading_dates:
        p = get_closing_price(sym, d.strftime('%Y-%m-%d'))
        if p:
            extra_masi[sym][d] = p
        time.sleep(0.15)
    print(f"  {sym} (extra MASI): {len(extra_masi[sym])} prices")

# Compute MASI day-by-day via chaining: MASI_t = MASI_{t-1} * (1 + Σ w_i * r_i)
sorted_dates = sorted(trading_dates)
masi_new = {}

for idx, d in enumerate(sorted_dates):
    prev_d = sorted_dates[idx - 1] if idx > 0 else None

    # Previous MASI value
    if idx == 0:
        prev_masi = float(prices[MASI_COL].dropna().iloc[-1])
    else:
        prev_masi = masi_new[prev_d]

    weighted_ret = 0.0
    coverage = 0.0

    for _, row in masi_w.iterrows():
        ticker = row['ticker']
        weight = float(row['weight'])
        if pd.isna(ticker):
            continue

        col = f'{ticker}.CS'

        # Today's price
        if col in new_data:
            today_p = new_data[col].get(d)
        else:
            today_p = extra_masi.get(ticker, {}).get(d)

        # Previous price — use prior new_data date if available, else last historical
        prev_p = None
        if prev_d is not None:
            if col in new_data:
                prev_p = new_data[col].get(prev_d)
            else:
                prev_p = extra_masi.get(ticker, {}).get(prev_d)
        if prev_p is None:
            if col in prices.columns:
                series = prices[col].dropna()
                prev_p = float(series.iloc[-1]) if len(series) > 0 else None

        if today_p and prev_p and prev_p > 0:
            weighted_ret += weight * ((today_p - prev_p) / prev_p)
            coverage += weight

    masi_new[d] = prev_masi * (1 + weighted_ret)
    print(f"  {d.date()}: MASI={masi_new[d]:.2f}  (prev={prev_masi:.2f}  wret={weighted_ret:+.4%}  cov={coverage:.1%})")

new_data[MASI_COL] = masi_new
print(f"  MASI: {len(masi_new)} values computed")

# ============================================================
# STEP 6 — Extend risk-free rate file for new trading days
# ============================================================
# The BDT 52-week rate (set at weekly BAM auctions) barely moves between auctions.
# We forward-fill the last known rate for any gap between the RF file and today.
print("\nUpdating risk-free rate file...")
RF_FILE = os.path.join(BASE_DIR, 'taux_sans_risque_maroc_quotidien.csv')
rf = pd.read_csv(RF_FILE, sep=';', parse_dates=['Date'], index_col='Date', encoding='utf-8-sig')
rf_last_date = rf.index.max()
rf_last_row  = rf.iloc[-1]

new_rf_dates = [d for d in sorted(trading_dates) if d > rf_last_date]
if new_rf_dates:
    new_rf_rows = []
    for d in new_rf_dates:
        row = rf_last_row.copy()
        row['Jour_Ouvre'] = True
        new_rf_rows.append(row)
    new_rf_df = pd.DataFrame(new_rf_rows, index=pd.DatetimeIndex(new_rf_dates))
    new_rf_df.index.name = 'Date'
    rf_extended = pd.concat([rf, new_rf_df]).sort_index()
    rf_extended.to_csv(RF_FILE, sep=';', date_format='%Y-%m-%d', encoding='utf-8-sig')
    print(f"  RF rate extended: {rf_last_date.date()} -> {max(new_rf_dates).date()} ({len(new_rf_dates)} days, rate={rf_last_row['BDT_52_semaines_%']}%)")
else:
    print(f"  RF rate already up to date ({rf_last_date.date()})")

# ============================================================
# STEP 7 — Build new rows DataFrame and append
# ============================================================
# Combine all new data into rows indexed by date
all_dates = sorted(trading_dates)
rows = []
for d in all_dates:
    row = {}
    for col in prices.columns:
        row[col] = new_data.get(col, {}).get(d, np.nan)
    rows.append(row)

new_df = pd.DataFrame(rows, index=pd.DatetimeIndex(all_dates))
new_df.index.name = 'Date'

print(f"\nNew rows built: {len(new_df)}")
print(f"  Range: {new_df.index.min().date()} to {new_df.index.max().date()}")
non_null = new_df.notna().sum().sum()
print(f"  Non-null cells: {non_null}/{len(new_df) * len(new_df.columns)}")

# Backup original
print(f"\nBacking up original...")
prices.to_csv(BACKUP_FILE, sep=';', date_format='%d/%m/%Y')

# Append
combined = pd.concat([prices, new_df]).sort_index()
combined.to_csv(PRICES_FILE, sep=';', date_format='%d/%m/%Y')

print(f"Saved: {PRICES_FILE}")
print(f"  Total rows: {len(combined)}")
print(f"  New end date: {combined.index.max().date()}")
print("\nDone. Run daily_automation.py to recompute all metrics.")
