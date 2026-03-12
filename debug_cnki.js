const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'cnki_auth.json' });
  const page = await context.newPage();
  
  console.log('正在打开知网高级搜索...');
  await page.goto('https://kns.cnki.net/kns8/defaultresult/index', { waitUntil: 'networkidle' });
  
  await page.waitForTimeout(3000);
  
  // 截图看页面结构
  await page.screenshot({ path: 'cnki_page.png', fullPage: true });
  console.log('页面截图已保存: cnki_page.png');
  
  // 获取页面内容
  const content = await page.content();
  console.log('页面HTML长度:', content.length);
  
  // 查找搜索框
  const inputs = await page.locator('input').all();
  console.log('找到', inputs.length, '个输入框');
  
  for (let i = 0; i < Math.min(inputs.length, 5); i++) {
    const placeholder = await inputs[i].getAttribute('placeholder').catch(() => '');
    const name = await inputs[i].getAttribute('name').catch(() => '');
    const id = await inputs[i].getAttribute('id').catch(() => '');
    console.log(`输入框 ${i}: placeholder=${placeholder}, name=${name}, id=${id}`);
  }
  
  await browser.close();
})();
