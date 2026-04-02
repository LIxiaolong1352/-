const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  try {
    // 访问密码验证页面
    await page.goto('https://www.wjx.cn/wjx/activitystat/verifyreportpassword.aspx?viewtype=1&activity=355467038&type=1', {
      waitUntil: 'networkidle'
    });
    
    // 输入密码
    await page.fill('#txtPassword', '123456');
    
    // 点击下一步
    await page.click('#btnContinue');
    
    // 等待页面加载
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(3000);
    
    // 获取页面内容
    const content = await page.content();
    console.log(content);
    
  } catch (error) {
    console.error('Error:', error.message);
  } finally {
    await browser.close();
  }
})();
