const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('搜索第三篇论文...');
  await page.goto('https://www.cnki.net', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(3000);
  
  // 找到搜索框并输入
  const allInputs = await page.locator('input').all();
  let searchInput = null;
  for (const input of allInputs) {
    const type = await input.getAttribute('type');
    if (type === 'text') {
      searchInput = input;
      break;
    }
  }
  
  if (searchInput) {
    await searchInput.fill('常见膝关节运动损伤与应对 张隽');
    console.log('已输入搜索关键词');
    
    // 查找并点击搜索按钮
    const buttons = await page.locator('button, .search-btn, [class*="search"]').all();
    for (const btn of buttons) {
      const text = await btn.textContent().catch(() => '');
      if (text.includes('搜索') || text.includes('检索') || text.includes('Search')) {
        await btn.click();
        console.log('点击搜索按钮');
        break;
      }
    }
  }
  
  await page.waitForTimeout(5000);
  
  // 截图搜索结果
  await page.screenshot({ path: 'cnki_paper3_search.png', fullPage: true });
  console.log('搜索结果截图: cnki_paper3_search.png');
  
  // 尝试点击第一篇结果
  const links = await page.locator('a').all();
  for (const link of links) {
    const text = await link.textContent().catch(() => '');
    if (text.includes('常见膝关节运动损伤与应对')) {
      await link.click();
      console.log('点击论文链接');
      break;
    }
  }
  
  await page.waitForTimeout(5000);
  await page.screenshot({ path: 'cnki_paper3_detail.png', fullPage: true });
  console.log('论文详情截图: cnki_paper3_detail.png');
  
  await browser.close();
  console.log('完成');
})();
