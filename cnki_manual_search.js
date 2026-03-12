const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(3000);
  
  // 直接操作搜索框
  console.log('查找搜索框...');
  const searchBox = await page.locator('input').filter({ hasText: '' }).first();
  
  // 获取所有可见的输入框
  const inputs = await page.locator('input:visible').all();
  console.log('找到', inputs.length, '个可见输入框');
  
  // 尝试第5个输入框（通常是主搜索框）
  if (inputs.length >= 5) {
    await inputs[4].fill('膝关节运动损伤');
    console.log('已输入关键词');
    
    // 查找搜索按钮
    const buttons = await page.locator('button:visible, .search-btn:visible, a:visible').all();
    console.log('找到', buttons.length, '个可见按钮/链接');
    
    // 截图
    await page.screenshot({ path: 'cnki_before_search.png' });
    console.log('截图保存: cnki_before_search.png');
  }
  
  console.log('请手动点击搜索按钮，然后按Enter继续...');
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  await page.waitForTimeout(3000);
  await page.screenshot({ path: 'cnki_after_search.png', fullPage: true });
  console.log('搜索结果截图: cnki_after_search.png');
  
  await browser.close();
})();
