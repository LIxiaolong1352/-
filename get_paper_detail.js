const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('打开第一篇论文...');
  // 第一篇论文的链接（根据标题搜索）
  await page.goto('https://kns.cnki.net/kns8/defaultresult/index', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(3000);
  
  // 搜索具体论文标题
  const inputs = await page.locator('input:visible').all();
  if (inputs.length > 0) {
    await inputs[inputs.length - 1].fill('常规康复训练联合NMES对膝关节韧带损伤重建后患者运动功能恢复的影响');
    
    // 点击搜索
    const searchBtn = await page.locator('.search-btn, button[type="submit"]').first();
    if (await searchBtn.isVisible()) {
      await searchBtn.click();
    }
  }
  
  await page.waitForTimeout(5000);
  await page.screenshot({ path: 'cnki_paper_detail.png', fullPage: true });
  console.log('论文详情页截图: cnki_paper_detail.png');
  
  await browser.close();
})();
