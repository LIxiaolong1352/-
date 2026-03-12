#!/usr/bin/env python3
"""币安价格邮件通知"""
import json
import urllib.request
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import sys

# 币安 API
API_KEY = "W1uAG2nfRxOiCritTTtbOYChO5tev9YMKLhWkEge2AGsmM3FNlk9g1kCDCYXaoW0"
PROXY = "http://127.0.0.1:7890"
BASE_URL = "https://api1.binance.com"

# 邮件配置
TO_EMAIL = "3796401811@qq.com"
SMTP_HOST = "smtp.qq.com"  # QQ邮箱
SMTP_PORT = 465

# 需要发件邮箱和密码/授权码
FROM_EMAIL = "3796401811@qq.com"  # 发件邮箱
FROM_PASSWORD = "tpugbxmctffucegj"  # QQ邮箱授权码

def get_price(symbol):
    url = f"{BASE_URL}/api/v3/ticker/price?symbol={symbol.upper()}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def get_24h_stats(symbol):
    url = f"{BASE_URL}/api/v3/ticker/24hr?symbol={symbol.upper()}"
    req = urllib.request.Request(url, headers={"X-MBX-APIKEY": API_KEY})
    req.set_proxy(PROXY.replace("http://", ""), "http")
    with urllib.request.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode())

def send_email(subject, content):
    if not FROM_EMAIL or not FROM_PASSWORD:
        print("错误: 请设置发件邮箱和授权码")
        return
    
    msg = MIMEText(content, 'plain', 'utf-8')
    msg['From'] = FROM_EMAIL
    msg['To'] = TO_EMAIL
    msg['Subject'] = Header(subject, 'utf-8')
    
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(FROM_EMAIL, FROM_PASSWORD)
            server.sendmail(FROM_EMAIL, [TO_EMAIL], msg.as_string())
        print(f"邮件已发送到 {TO_EMAIL}")
    except Exception as e:
        print(f"邮件发送失败: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python binance_email.py <币种>")
        print("示例: python binance_email.py BTCUSDT")
        sys.exit(1)
    
    symbol = sys.argv[1]
    try:
        price_data = get_price(symbol)
        stats = get_24h_stats(symbol)
        
        price = float(price_data['price'])
        change = float(stats['priceChangePercent'])
        high = float(stats['highPrice'])
        low = float(stats['lowPrice'])
        
        subject = f"{symbol} 价格提醒: ${price:,.2f} ({change:+.2f}%)"
        content = f"""币安 {symbol} 价格更新

当前价格: ${price:,.2f}
24h涨跌: {change:+.2f}%
24h最高: ${high:,.2f}
24h最低: ${low:,.2f}
24h成交量: {float(stats['volume']):,.2f}

发送时间: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        print(content)
        send_email(subject, content)
        
    except Exception as e:
        print(f"错误: {e}")
