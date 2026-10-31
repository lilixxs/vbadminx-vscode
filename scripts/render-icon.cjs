const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');

async function renderIcon() {
  const svgPath = path.resolve('resources/icon.svg');
  const pngPath = path.resolve('resources/icon.png');
  const svgContent = fs.readFileSync(svgPath, 'utf8');

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({
    viewport: { width: 256, height: 256 },
    deviceScaleFactor: 1
  });

  const html = `<!DOCTYPE html>
<html>
<head>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { background: transparent; overflow: hidden; display: flex; }
  </style>
</head>
<body>
  ${svgContent}
</body>
</html>`;

  await page.setContent(html);
  const buffer = await page.screenshot({
    omitBackground: true,
    type: 'png',
    clip: { x: 0, y: 0, width: 256, height: 256 }
  });

  fs.writeFileSync(pngPath, buffer);
  await browser.close();
  console.log(`Successfully generated ${pngPath} (${buffer.length} bytes)`);
}

renderIcon().catch(err => {
  console.error(err);
  process.exit(1);
});
