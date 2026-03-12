const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('正在打开知网...');
  await page.goto('https://kns.cnki.net/kns8/AdvSearch', { waitUntil: 'networkidle' });
  await page.waitForTimeout(5000);
  
  // 截图看页面
  await page.screenshot({ path: 'cnki_adv_search.png', fullPage: true });
  console.log('高级搜索页面截图: cnki_adv_search.png');
  
  // 尝试多种方式找输入框
  const selectors = [
    'input[type="text"]',
    '.search-input',
    '#SearchWord',
    'input[name*="keyword"]',
    'input[name*="search"]'
  ];
  
  for (const selector of selectors) {
    const count = await page.locator(selector).count();
    if (count > 0) {
      console.log(`找到 ${count} 个 ${selector}`);
    }
  }
  
  await browser.close();
})();
