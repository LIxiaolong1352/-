const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('正在打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'networkidle' });
  await page.waitForTimeout(3000);
  
  // 查找搜索框 - 使用placeholder
  const searchInput = page.locator('input[placeholder*="文献"]').first();
  await searchInput.waitFor({ timeout: 10000 });
  
  console.log('输入搜索关键词: 膝关节运动损伤');
  await searchInput.fill('膝关节运动损伤');
  
  // 点击搜索按钮
  const searchBtn = page.locator('.search-btn, .btn-search, button:has-text("检索"), .search-icon').first();
  await searchBtn.click();
  
  console.log('等待搜索结果...');
  await page.waitForTimeout(5000);
  
  // 截图
  await page.screenshot({ path: 'cnki_search.png', fullPage: true });
  console.log('搜索结果截图: cnki_search.png');
  
  // 获取当前URL
  console.log('当前URL:', page.url());
  
  // 尝试获取论文列表
  const articles = await page.locator('.result-table-list tbody tr, .article-list .item, .result-list .item').all();
  console.log('找到', articles.length, '条结果');
  
  if (articles.length > 0) {
    for (let i = 0; i < Math.min(articles.length, 3); i++) {
      const title = await articles[i].locator('.title a, .article-title, a[title]').textContent().catch(() => '无标题');
      console.log(`\n论文 ${i + 1}:`, title?.trim());
    }
  }
  
  await browser.close();
  console.log('完成');
})();
