const fs = require('node:fs');
const path = require('node:path');
const { parseFormDesign } = require('../dist/src/formDesign');
const { renderFormPage } = require('../dist/src/formHtml');
const { chromium } = require('playwright');

/**
 * 生成示例窗体的预览截图（开发/文档用途，不进入扩展包）。
 * 使用 examples/LoginForm.frm → output/form-preview.html → output/screenshots/09-form-preview.png
 */
async function main() {
  const source = path.join(__dirname, '..', 'examples', 'LoginForm.frm');
  const output = path.join(__dirname, '..', 'output');
  const html = path.join(output, 'form-preview.html');
  const screenshots = path.join(output, 'screenshots');
  fs.mkdirSync(screenshots, { recursive: true });

  const page = renderFormPage(parseFormDesign(fs.readFileSync(source, 'utf8')), {
    nonce: 'render-preview',
    title: 'Form preview: LoginForm.frm',
    hint: 'VBA UserForm preview (approximate rendering)',
    emptyText: 'No design',
  });
  fs.writeFileSync(html, page);

  const browser = await chromium.launch();
  try {
    const view = await browser.newPage({ viewport: { width: 960, height: 760 }, deviceScaleFactor: 1 });
    await view.goto(`file:///${html.replaceAll('\\', '/')}`);
    await view.waitForTimeout(250);
    await view.screenshot({ path: path.join(screenshots, '09-form-preview.png'), fullPage: true });
  } finally {
    await browser.close();
  }
  console.log('wrote output/screenshots/09-form-preview.png');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
