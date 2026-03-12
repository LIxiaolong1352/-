const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'load' });
  await page.waitForTimeout(3000);
  
  console.log('请手动操作：');
  console.log('1. 搜索：常见膝关节运动损伤与应对');
  console.log('2. 点击论文标题进入详情页（不是勾选框，是标题文字）');
  console.log('3. 进入详情页后按 Enter 键');
  
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  await page.waitForTimeout(2000);
  await page.screenshot({ path: 'paper3_detail.png', fullPage: true });
  console.log('已截图');
  
  await browser.close();
})();
