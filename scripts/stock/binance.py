#!/usr/bin/env python3
"""币安 API 数据源 - 获取实时价格"""
import json
import urllib.request
import urllib.parse
import hmac
import hashlib
import time

API_KEY = "W1uAG2nfRxOiCritTTtbOYChO5tev9YMKLhWkEge2AGsmM3FNlk9g1kCDCYXaoW0"
API_SECRET = ""
BASE_URL = "https://api1.binance.com"  # 备用域名

# Clash 代理
PROXY = "http://127.0.0.1:7890"

def get_price(symbol):
    """获取单个币种价格"""
    url = f"{BASE_URL}/api/v3/ticker/price?symbol={symbol.upper()}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def get_24h_stats(symbol):
    """获取24小时统计"""
    url = f"{BASE_URL}/api/v3/ticker/24hr?symbol={symbol.upper()}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def get_all_prices():
    """获取所有币种价格"""
    url = f"{BASE_URL}/api/v3/ticker/price"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python binance.py <币种>")
        print("示例: python binance.py BTCUSDT")
        sys.exit(1)
    
    symbol = sys.argv[1]
    try:
        price_data = get_price(symbol)
        stats = get_24h_stats(symbol)
        
        print(f"\n[币安 {symbol}]")
        print(f"当前价格: ${float(price_data['price']):,.2f}")
        print(f"24h涨跌: {float(stats['priceChangePercent']):+.2f}%")
        print(f"24h最高: ${float(stats['highPrice']):,.2f}")
        print(f"24h最低: ${float(stats['lowPrice']):,.2f}")
        print(f"24h成交量: {float(stats['volume']):,.2f}")
    except Exception as e:
        print(f"错误: {e}")
