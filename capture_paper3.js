const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('请在当前页面点击第三篇论文《常见膝关节运动损伤与应对》');
  console.log('进入详情页后按 Enter 键，我帮你截图分析...');
  
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  await page.waitForTimeout(3000);
  await page.screenshot({ path: 'cnki_paper3_detail.png', fullPage: true });
  console.log('论文详情截图: cnki_paper3_detail.png');
  
  await browser.close();
  console.log('完成');
})();
