#!/usr/bin/env python3
"""扫描低价山寨币+大资金流入"""
import json
import urllib.request

API_KEY = "W1uAG2nfRxOiCritTTtbOYChO5tev9YMKLhWkEge2AGsmM3FNlk9g1kCDCYXaoW0"
PROXY = "http://127.0.0.1:7890"
BASE_URL = "https://api1.binance.com"

def get_all_24h():
    """获取所有交易对24h数据"""
    url = f"{BASE_URL}/api/v3/ticker/24hr"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())

# 获取数据
data = get_all_24h()

# 筛选条件
results = []
for item in data:
    symbol = item['symbol']
    # 只选USDT交易对，排除杠杆代币
    if not symbol.endswith('USDT') or 'UP' in symbol or 'DOWN' in symbol or 'BEAR' in symbol or 'BULL' in symbol:
        continue
    
    price = float(item['lastPrice'])
    change = float(item['priceChangePercent'])
    volume = float(item['volume'])
    quote_volume = float(item['quoteVolume'])  # USDT成交量
    
    # 条件：价格<5刀，成交量>1000万USDT，涨幅>0（有资金流入）
    if price < 5 and quote_volume > 10000000 and change > 0:
        results.append({
            'symbol': symbol,
            'price': price,
            'change': change,
            'volume_usdt': quote_volume,
            'volume_coin': volume
        })

# 按成交量排序
results.sort(key=lambda x: x['volume_usdt'], reverse=True)

# 打印前20
print("[低价山寨币 + 大资金流入]\n")
print(f"{'排名':<4} {'币种':<12} {'价格':<10} {'24h涨跌':<10} {'成交量(USDT)':<15}")
print("-" * 55)
for i, r in enumerate(results[:20], 1):
    print(f"{i:<4} {r['symbol']:<12} ${r['price']:<9.4f} {r['change']:>+7.2f}%    ${r['volume_usdt']:>12,.0f}")

print(f"\n共找到 {len(results)} 个符合条件的币种")
