// Render a BOV deck HTML file to PDF via headless Chromium (Playwright).
// Usage: NODE_PATH=<global node_modules> node render.js <input.html> <output.pdf>
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const [input, output] = process.argv.slice(2);
  if (!input || !output) { console.error('usage: node render.js <input.html> <output.pdf>'); process.exit(1); }
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('file://' + path.resolve(input), { waitUntil: 'networkidle' });
  await page.pdf({ path: output, width: '1700px', height: '1080px',
    printBackground: true, margin: { top: 0, bottom: 0, left: 0, right: 0 } });
  await browser.close();
  console.log('rendered:', output);
})();
