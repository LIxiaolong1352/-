const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('正在打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'networkidle' });
  
  // 等待搜索框加载
  await page.waitForSelector('#txt_SearchText', { timeout: 10000 });
  
  // 输入搜索关键词
  console.log('搜索: 膝关节运动损伤');
  await page.fill('#txt_SearchText', '膝关节运动损伤');
  await page.click('.search-btn');
  
  // 等待搜索结果
  await page.waitForTimeout(3000);
  
  // 设置筛选条件 - 近一年
  console.log('设置时间筛选: 近一年');
  await page.click('text=发表时间');
  await page.waitForTimeout(1000);
  
  // 选择近一年
  const yearFilter = await page.locator('text=近一年').first();
  if (await yearFilter.isVisible()) {
    await yearFilter.click();
    await page.waitForTimeout(2000);
  }
  
  // 获取第一篇论文信息
  console.log('获取第一篇论文信息...');
  const firstArticle = await page.locator('.result-table-list .item').first();
  
  if (await firstArticle.isVisible()) {
    const title = await firstArticle.locator('.title a').textContent();
    const authors = await firstArticle.locator('.author').textContent();
    const source = await firstArticle.locator('.source').textContent();
    const date = await firstArticle.locator('.date').textContent();
    
    console.log('\n=== 第一篇论文 ===');
    console.log('标题:', title?.trim());
    console.log('作者:', authors?.trim());
    console.log('来源:', source?.trim());
    console.log('日期:', date?.trim());
    
    // 截图保存搜索结果
    await page.screenshot({ path: 'cnki_search_results.png', fullPage: true });
    console.log('\n截图已保存: cnki_search_results.png');
  } else {
    console.log('未找到论文');
  }
  
  await browser.close();
  console.log('完成');
})();
