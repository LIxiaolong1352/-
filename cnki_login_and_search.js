const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  
  console.log('正在打开知网登录页面...');
  await page.goto('https://www.cnki.net', { waitUntil: 'networkidle' });
  
  console.log('页面已打开，请登录你的知网账号');
  console.log('登录完成后按 Enter 键保存登录状态...');
  
  // 等待用户按回车
  process.stdin.setRawMode(true);
  process.stdin.resume();
  await new Promise(resolve => process.stdin.once('data', resolve));
  
  // 保存登录状态
  await context.storageState({ path: 'cnki_auth.json' });
  console.log('登录状态已保存！');
  
  // 测试搜索
  console.log('测试搜索: 膝关节运动损伤');
  await page.goto('https://kns.cnki.net/kns8/defaultresult/index?kw=%E8%86%9D%E5%85%B3%E8%8A%82%E8%BF%90%E5%8A%A8%E6%8D%9F%E4%BC%A4&korder=SU', { waitUntil: 'networkidle' });
  await page.waitForTimeout(3000);
  
  // 截图
  await page.screenshot({ path: 'cnki_search_result.png', fullPage: true });
  console.log('搜索结果已截图保存');
  
  await browser.close();
  console.log('完成');
})();
