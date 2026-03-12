const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  
  console.log('正在打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'networkidle' });
  
  console.log('页面已加载，请手动登录');
  console.log('登录完成后按 Enter 键继续...');
  
  // 等待用户按回车
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  // 保存登录状态
  await context.storageState({ path: 'cnki_auth.json' });
  console.log('登录状态已保存: cnki_auth.json');
  
  await browser.close();
  console.log('完成');
})();
