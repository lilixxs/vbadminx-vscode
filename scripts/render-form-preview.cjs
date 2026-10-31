const fs = require('node:fs');
const path = require('node:path');
const { parseFormDesign } = require('../dist/src/formDesign');
const { inferFormDesignFromCode } = require('../dist/src/formInfer');
const { renderFormPage } = require('../dist/src/formHtml');
const { chromium } = require('playwright');

/**
 * 生成示例窗体的预览截图（开发/文档用途，不进入扩展包）。
 * - examples/LoginForm.frm：完整设计段 → 09-form-preview.png
 * - examples/CodeOnlyForm.frm：仅代码 → 10-form-inferred.png（推断预览）
 */
function analyzed(source, name) {
  const text = fs.readFileSync(source, 'utf8');
  const parsed = parseFormDesign(text);
  if (parsed.root) { return { model: parsed, notice: '' }; }
  const inferred = inferFormDesignFromCode(text, name);
  return { model: inferred, notice: inferred.root ? '此 .frm 没有窗体设计段。预览由代码中的布局赋值与事件过程推断，可能与实际窗体不同。' : '' };
}

async function main() {
  const output = path.join(__dirname, '..', 'output');
  const screenshots = path.join(output, 'screenshots');
  fs.mkdirSync(screenshots, { recursive: true });

  const jobs = [
    { source: path.join(__dirname, '..', 'examples', 'LoginForm.frm'), name: 'LoginForm', out: '09-form-preview.png' },
    { source: path.join(__dirname, '..', 'examples', 'CodeOnlyForm.frm'), name: 'CodeOnlyForm', out: '10-form-inferred.png' },
  ];

  const pages = [];
  for (const job of jobs) {
    const { model, notice } = analyzed(job.source, job.name);
    const title = `Form preview: ${path.basename(job.source)}`;
    const html = renderFormPage(model, {
      nonce: `preview-${job.name}`,
      title,
      hint: 'VBA UserForm preview (approximate rendering)',
      emptyText: 'No design',
      notice,
    });
    const file = path.join(output, `form-${job.name}.html`);
    fs.writeFileSync(file, html);
    pages.push({ file, out: path.join(screenshots, job.out) });
  }

  const browser = await chromium.launch();
  try {
    for (const page of pages) {
      const view = await browser.newPage({ viewport: { width: 960, height: 760 }, deviceScaleFactor: 1 });
      await view.goto(`file:///${page.file.replaceAll('\\', '/')}`);
      await view.waitForTimeout(250);
      await view.screenshot({ path: page.out, fullPage: true });
      await view.close();
    }
  } finally {
    await browser.close();
  }
  console.log('wrote output/screenshots/09-form-preview.png and 10-form-inferred.png');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
