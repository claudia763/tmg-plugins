const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1530, height: 1000 } });
  await page.goto('file://' + path.resolve('map.html'));
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(6000);
  await page.screenshot({ path: 'comp_map.png' });
  await browser.close();
  console.log('wrote comp_map.png');
})();
