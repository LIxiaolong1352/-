const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  try {
    // 搜索力合科技股票代码
    await page.goto('https://www.google.com/search?q=力合科技+股票代码+深圳', { waitUntil: 'networkidle' });
    await page.waitForTimeout(2000);
    
    // 获取页面文本内容
    const content = await page.locator('body').textContent();
    console.log('=== Google 搜索结果 ===');
    console.log(content.substring(0, 3000));
    
    // 尝试查找股票代码
    const stockMatch = content.match(/(\d{6}\.SZ|\d{6}\.SS|300\d{3}|002\d{3}|000\d{3})/);
    if (stockMatch) {
      console.log('\n=== 找到的股票代码 ===');
      console.log(stockMatch[0]);
    }
    
  } catch (e) {
    console.error('Error:', e.message);
  }
  
  await browser.close();
})();
