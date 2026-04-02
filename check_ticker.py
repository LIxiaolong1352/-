import yfinance as yf

# 可能的力合科技代码
tickers = ['300800.SZ', '002243.SZ', '832885.BJ', '300425.SZ']

for t in tickers:
    try:
        info = yf.Ticker(t).info
        name = info.get('longName', 'N/A')
        symbol = info.get('symbol', 'N/A')
        print(f'{t} -> {symbol}: {name}')
    except Exception as e:
        print(f'{t}: Error - {e}')
