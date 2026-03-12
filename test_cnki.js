const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  
  console.log('正在打开知网...');
  await page.goto('https://www.cnki.net', { waitUntil: 'networkidle' });
  
  console.log('页面标题:', await page.title());
  console.log('页面URL:', page.url());
  
  // 截图保存
  await page.screenshot({ path: 'cnki_homepage.png', fullPage: true });
  console.log('截图已保存: cnki_homepage.png');
  
  // 等待30秒让你手动登录
  console.log('请在30秒内手动登录知网...');
  await page.waitForTimeout(30000);
  
  // 保存登录状态
  await context.storageState({ path: 'cnki_auth.json' });
  console.log('登录状态已保存: cnki_auth.json');
  
  await browser.close();
  console.log('测试完成');
})();
