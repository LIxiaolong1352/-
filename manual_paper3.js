const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('重新打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'load' });
  await page.waitForTimeout(5000);
  
  console.log('页面已加载');
  await page.screenshot({ path: 'cnki_current.png' });
  console.log('当前页面截图: cnki_current.png');
  
  console.log('\n请手动操作：');
  console.log('1. 在搜索框输入：常见膝关节运动损伤与应对');
  console.log('2. 点击搜索');
  console.log('3. 点击论文标题进入详情页');
  console.log('4. 完成后按 Enter 键');
  
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  await page.waitForTimeout(2000);
  await page.screenshot({ path: 'cnki_paper3_final.png', fullPage: true });
  console.log('论文详情已截图: cnki_paper3_final.png');
  
  await browser.close();
})();
