#!/usr/bin/env python3
"""XRP 深度分析"""
import json
import urllib.request

API_KEY = "W1uAG2nfRxOiCritTTtbOYChO5tev9YMKLhWkEge2AGsmM3FNlk9g1kCDCYXaoW0"
PROXY = "http://127.0.0.1:7890"
BASE_URL = "https://api1.binance.com"

def get_24h(symbol):
    url = f"{BASE_URL}/api/v3/ticker/24hr?symbol={symbol.upper()}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def get_order_book(symbol, limit=100):
    """获取订单簿"""
    url = f"{BASE_URL}/api/v3/depth?symbol={symbol.upper()}&limit={limit}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def get_recent_trades(symbol, limit=100):
    """获取最近成交"""
    url = f"{BASE_URL}/api/v3/trades?symbol={symbol.upper()}&limit={limit}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

symbol = "XRPUSDT"
data = get_24h(symbol)

# 基础数据
price = float(data['lastPrice'])
change = float(data['priceChangePercent'])
high = float(data['highPrice'])
low = float(data['lowPrice'])
volume = float(data['volume'])
quote_volume = float(data['quoteVolume'])
weighted_avg = float(data['weightedAvgPrice'])

# 计算波动率
volatility = ((high - low) / low) * 100

# 买盘卖盘压力
order_book = get_order_book(symbol, 50)
bid_volume = sum([float(b[1]) for b in order_book['bids']])
ask_volume = sum([float(a[1]) for a in order_book['asks']])
buy_pressure = (bid_volume / (bid_volume + ask_volume)) * 100

# 大单分析（最近100笔）
trades = get_recent_trades(symbol, 100)
big_trades = [t for t in trades if float(t['quoteQty']) > 10000]  # 大于1万USDT
big_buy = len([t for t in big_trades if not t['isBuyerMaker']])
big_sell = len([t for t in big_trades if t['isBuyerMaker']])

print(f"""
[XRPUSDT 深度分析]

【价格数据】
当前价格: ${price:.4f}
24h涨跌: {change:+.2f}%
24h最高: ${high:.4f}
24h最低: ${low:.4f}
加权均价: ${weighted_avg:.4f}
波动率: {volatility:.2f}%

【成交数据】
24h成交量: {volume:,.0f} XRP
24h成交额: ${quote_volume:,.0f} USDT

【买卖压力】
买盘挂单: {bid_volume:,.0f} XRP
卖盘挂单: {ask_volume:,.0f} XRP
买方压力: {buy_pressure:.1f}%

【大单分析】(最近100笔中>${10000:,} USDT)
大单总数: {len(big_trades)}
大单买入: {big_buy}
大单卖出: {big_sell}
大单净买入: {big_buy - big_sell:+d}

【评估】
{"强势买入信号!" if buy_pressure > 55 and big_buy > big_sell else "买入占优" if buy_pressure > 50 else "卖出占优" if buy_pressure < 50 else "均衡"}
""")
