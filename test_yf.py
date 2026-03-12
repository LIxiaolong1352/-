import yfinance as yf
data = yf.Ticker('AAPL').info
print(f"AAPL: ${data.get('currentPrice', 'N/A')}")