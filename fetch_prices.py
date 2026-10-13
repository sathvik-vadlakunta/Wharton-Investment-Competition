import yfinance as yf
import json, os, time, openpyxl
from datetime import datetime

OUT = os.path.join(os.path.dirname(__file__), 'data', 'prices')
os.makedirs(OUT, exist_ok=True)

# Load tickers from Excel
wb = openpyxl.load_workbook('/Users/sathvikvadlakunta/Downloads/WGHIC_Approved_Securities_Clean.xlsx')

stocks = [(r[1], r[0]) for r in wb['Stocks'].iter_rows(min_row=2, values_only=True) if r[1]]
etfs   = [(r[1], r[2]) for r in wb['ETFs'].iter_rows(min_row=2, values_only=True) if r[1]]

all_tickers = stocks + etfs
print(f"Total tickers: {len(all_tickers)}")

failed = []
done = 0

# Batch download in groups of 50
BATCH = 50
ticker_list = [t[0] for t in all_tickers]

for i in range(0, len(ticker_list), BATCH):
    batch = ticker_list[i:i+BATCH]
    print(f"Batch {i//BATCH + 1}: downloading {len(batch)} tickers...")
    try:
        raw = yf.download(batch, period='2y', interval='1d', auto_adjust=True, progress=False)
        for ticker in batch:
            out_path = os.path.join(OUT, f"{ticker}.json")
            if os.path.exists(out_path):
                done += 1
                continue
            try:
                if len(batch) == 1:
                    df = raw
                else:
                    df = raw.xs(ticker, axis=1, level=1) if ticker in raw.columns.get_level_values(1) else None
                if df is None or df.empty:
                    failed.append(ticker)
                    continue
                df = df.dropna(subset=['Close'])
                records = []
                for date, row in df.iterrows():
                    records.append({
                        'd': date.strftime('%Y-%m-%d'),
                        'o': round(float(row['Open']), 4) if not hasattr(row['Open'], '__iter__') else None,
                        'h': round(float(row['High']), 4) if not hasattr(row['High'], '__iter__') else None,
                        'l': round(float(row['Low']), 4) if not hasattr(row['Low'], '__iter__') else None,
                        'c': round(float(row['Close']), 4) if not hasattr(row['Close'], '__iter__') else None,
                        'v': int(row['Volume']) if not hasattr(row['Volume'], '__iter__') else 0,
                    })
                records = [r for r in records if r['c'] is not None]
                with open(out_path, 'w') as f:
                    json.dump(records, f, separators=(',', ':'))
                done += 1
            except Exception as e:
                failed.append(ticker)
                print(f"  failed {ticker}: {e}")
    except Exception as e:
        print(f"  batch error: {e}")
        failed += batch
    time.sleep(1)

print(f"\nDone: {done} saved, {len(failed)} failed")
if failed:
    print("Failed:", failed[:30])
