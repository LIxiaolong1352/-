#!/usr/bin/env python3
"""
CoinGlass 虚拟币市场数据爬取工具
使用时运行，不需要 API 密钥
"""

import subprocess
import json
from datetime import datetime

def scrape_coinglass():
    """爬取 CoinGlass 关键数据"""
    
    js_code = '''
const { chromium } = require('playwright');

(async () => {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    
    try {
        // 访问 CoinGlass
        await page.goto('https://www.coinglass.com/', { 
            waitUntil: 'domcontentloaded',
            timeout: 20000 
        });
        
        // 等待页面加载
        await page.waitForTimeout(5000);
        
        // 获取主要币种价格数据
        const data = await page.evaluate(() => {
            const result = {
                timestamp: new Date().toISOString(),
                btc: {},
                eth: {},
                market_summary: []
            };
            
            // 尝试获取表格数据
            const rows = document.querySelectorAll('table tr');
            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                if (cells.length > 3) {
                    const symbol = cells[0]?.textContent?.trim();
                    if (symbol && (symbol.includes('BTC') || symbol.includes('ETH'))) {
                        result.market_summary.push({
                            symbol: symbol,
                            price: cells[1]?.textContent?.trim(),
                            change_24h: cells[2]?.textContent?.trim(),
                            volume: cells[3]?.textContent?.trim()
                        });
                    }
                }
            });
            
            return result;
        });
        
        console.log(JSON.stringify(data, null, 2));
        
    } catch (error) {
        console.log(JSON.stringify({ error: error.message }));
    }
    
    await browser.close();
})();
'''
    
    # 执行 Node.js 代码
    result = subprocess.run(
        ['node', '-e', js_code],
        capture_output=True,
        text=True,
        timeout=60
    )
    
    try:
        return json.loads(result.stdout)
    except:
        return { 'error': '解析失败', 'raw': result.stdout }

def analyze_market(data):
    """简单分析市场数据"""
    analysis = {
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'raw_data': data,
        'summary': '数据获取完成',
        'notes': [
            '建议关注：爆仓热力图',
            '建议关注：资金费率',
            '建议关注：持仓量变化',
            '建议结合技术分析判断趋势'
        ]
    }
    return analysis

if __name__ == '__main__':
    print('正在爬取 CoinGlass 数据...')
    data = scrape_coinglass()
    analysis = analyze_market(data)
    print(json.dumps(analysis, indent=2, ensure_ascii=False))
