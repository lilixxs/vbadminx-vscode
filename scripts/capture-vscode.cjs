const { chromium } = require('playwright');
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

async function main() {
  const root = path.resolve(__dirname, '..');
  const output = path.join(root, 'output', 'screenshots');
  const callerSource = fs.readFileSync(path.join(root, 'examples', 'Caller.bas'), 'utf8');
  fs.mkdirSync(output, { recursive: true });
  const executable = process.env.VSCODE_EXECUTABLE_PATH || path.join(process.env.LOCALAPPDATA, 'Programs', 'Microsoft VS Code', 'Code.exe');
  const port = 19347;
  const installedExtensions = process.env.VBA_INSTALLED_EXTENSIONS;
  const profile = fs.mkdtempSync(path.join(root, '.vscode-test', 'capture-'));
  fs.mkdirSync(path.join(profile, 'User'), { recursive: true });
  fs.writeFileSync(path.join(profile, 'User', 'settings.json'), JSON.stringify({
    'workbench.startupEditor': 'none', 'window.restoreWindows': 'none', 'chat.disableAIFeatures': true,
    'editor.fontSize': 16, 'window.zoomLevel': 0, 'workbench.colorTheme': 'Default Dark Modern',
    'files.autoSave': 'off', 'editor.quickSuggestions': { other: 'on', comments: 'off', strings: 'off' },
  }));
  const child = spawn(executable, [
    `--remote-debugging-port=${port}`, '--disable-workspace-trust', '--skip-welcome', '--skip-release-notes',
    `--user-data-dir=${profile}`,
    ...(installedExtensions ? [`--extensions-dir=${installedExtensions}`] : ['--disable-extensions', `--extensionDevelopmentPath=${root}`]),
    path.join(root, 'examples'), path.join(root, 'examples', 'Caller.bas'),
  ], { stdio: 'ignore' });
  let browser;
  try {
    const stop = Date.now() + 30000;
    while (Date.now() < stop) {
      try { browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`); break; }
      catch { await new Promise(resolve => setTimeout(resolve, 500)); }
    }
    if (!browser) throw new Error('未能连接 VS Code 调试端口。');
    const context = browser.contexts()[0];
    const page = context.pages().find(page => /workbench/.test(page.url())) || context.pages()[0];
    await page.waitForSelector('.monaco-editor', { timeout: 30000 });
    await page.locator('.monaco-workbench').waitFor();
    await page.waitForFunction(() => document.body.innerText.includes('VBAdminx'), { timeout: 30000 });
    await page.screenshot({ path: path.join(output, '01-language-editor.png') });
    await page.locator('.monaco-editor .view-lines').first().click({ position: { x: 250, y: 100 } });
    await page.bringToFront();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    await page.keyboard.type('    result = AddN');
    await page.keyboard.press('Control+Space');
    await page.waitForSelector('.suggest-widget.visible', { timeout: 10000 }).catch(async () => {
      await page.keyboard.press('Control+Shift+p');
      await page.keyboard.type('Trigger Suggest');
      await page.keyboard.press('Enter');
      try { await page.waitForSelector('.suggest-widget.visible', { timeout: 10000 }); }
      catch (error) {
        await page.screenshot({ path: path.join(output, 'debug-suggest.png') });
        console.log((await page.locator('body').innerText()).slice(-4000));
        throw error;
      }
    });
    await page.screenshot({ path: path.join(output, '02-completion.png') });
    await page.keyboard.press('Escape');
    await page.keyboard.press('Control+a');
    await page.keyboard.insertText(callerSource);
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Control+Shift+p');
    await page.keyboard.type('VBAdminx:');
    await page.waitForSelector('.quick-input-widget', { timeout: 10000 });
    await page.screenshot({ path: path.join(output, '03-commands.png') });
    await page.keyboard.press('Escape');
    await page.locator('.monaco-editor .view-lines').first().click({ position: { x: 250, y: 100 } });
    await page.keyboard.press('Control+a');
    await page.keyboard.insertText(callerSource.replace('    result =', '    undeclaredValue = 1\n    result ='));
    await page.keyboard.press('Control+Shift+m');
    await page.waitForFunction(() => {
      const problems = document.querySelector('.markers-panel');
      return problems && problems.textContent.includes('undeclaredValue');
    }, { timeout: 10000 });
    await page.screenshot({ path: path.join(output, '04-diagnostics.png') });

    // 05 悬停提示：恢复示例源码，光标移到 AddNumbers 上，用 Ctrl+K Ctrl+I 触发 Show Hover。
    await page.locator('.monaco-editor .view-lines').first().click({ position: { x: 250, y: 100 } });
    await page.keyboard.press('Control+a');
    await page.keyboard.insertText(callerSource);
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < 4; i += 1) { await page.keyboard.press('ArrowDown'); }
    await page.keyboard.press('Home');
    for (let i = 0; i < 14; i += 1) { await page.keyboard.press('ArrowRight'); }
    await page.keyboard.press('Control+k');
    await page.keyboard.press('Control+i');
    await page.waitForFunction(() => {
      const hover = document.querySelector('.monaco-hover, .hover-content, .monaco-hover-content');
      return !!hover && hover.textContent.includes('AddNumbers');
    }, { timeout: 10000 });
    await page.screenshot({ path: path.join(output, '05-hover.png') });
    await page.keyboard.press('Escape');

    // 06 工程索引：运行“查看工程索引”命令，展示服务器返回的工程状态 JSON。
    await page.keyboard.press('Control+Shift+p');
    await page.keyboard.type('VBAdminx: 查看工程索引');
    await page.waitForSelector('.quick-input-widget', { timeout: 10000 });
    await page.keyboard.press('Enter');
    await page.waitForFunction(() => document.body.innerText.includes('"projects"'), { timeout: 10000 });
    await page.screenshot({ path: path.join(output, '06-project-status.png') });

    // 07 符号面板：回到 VBA 示例 → 用标题栏按钮切换为“整个工程” → 聚焦视图并截图。
    await page.keyboard.press('Control+p');
    await page.waitForSelector('.quick-input-widget', { timeout: 10000 });
    await page.keyboard.type('Caller.bas');
    await page.waitForTimeout(600);
    await page.keyboard.press('Enter');
    await page.waitForFunction(() => document.body.innerText.includes('Public Sub Main'), { timeout: 20000 });

    // 标题栏“符号视图选项”按钮：三步快捷菜单（范围 / 排序 / 过滤）。
    // 一次命令 = 三步菜单（范围 → 排序 → 过滤），必须连续完成，不能重复调用命令。
    const choose = async (text) => {
      await page.waitForSelector('.quick-input-widget .monaco-list-row', { timeout: 10000 });
      await page.keyboard.type(text);
      await page.waitForTimeout(400);
      await page.keyboard.press('Enter');
    };
    await page.keyboard.press('Control+Shift+p');
    await page.keyboard.type('VBAdminx: 符号视图选项');
    await page.waitForSelector('.quick-input-widget', { timeout: 10000 });
    await page.keyboard.press('Enter');
    await choose('整个工程');
    await choose('按出现位置');
    await choose('全部');

    await page.keyboard.press('Control+Shift+p');
    await page.keyboard.type('VBAdminx: 聚焦符号视图');
    await page.waitForSelector('.quick-input-widget', { timeout: 10000 });
    await page.keyboard.press('Enter');
    // 必须检查“树内”是否列出符号，不能只检查页面文本（编辑器里也有同名文本）。
    const rows = await page.locator('.monaco-list-row').allTextContents();
    console.log('TREE ROWS:', JSON.stringify(rows.slice(0, 40)));
    try {
      await page.waitForFunction(() => {
        const rows = [...document.querySelectorAll('.monaco-list-row')];
        return rows.some(row => row.textContent.includes('AddNumbers'));
      }, { timeout: 20000 });
    } catch (error) {
      await page.screenshot({ path: path.join(output, 'debug-symbols.png') });
      console.log('SYMBOLS DEBUG ROWS:', JSON.stringify((await page.locator('.monaco-list-row').allTextContents()).slice(0, 40)));
      throw error;
    }
    await page.screenshot({ path: path.join(output, '07-symbols-view.png') });

    // 08 点击树项 AddNumbers → 编辑器跳转并选中该名称，左侧同步显示。
    // 只在侧边栏树行中点击，避免命中编辑器里的同名文本。
    await page.locator('.monaco-list-row').filter({ hasText: 'AddNumbers' }).first().click();
    try {
      // 跳转证据：切换到 Demo.bas 且状态栏出现选中计数（AddNumbers 为 10 个字符）。
      await page.waitForFunction(() => {
        const text = document.body.innerText;
        return text.includes('Demo.bas') && /selected/i.test(text);
      }, { timeout: 15000 });
    } catch (error) {
      await page.screenshot({ path: path.join(output, 'debug-symbols.png') });
      console.log('JUMP DEBUG:', (await page.locator('body').innerText()).slice(0, 1500));
      throw error;
    }
    await page.screenshot({ path: path.join(output, '08-symbols-jump.png') });

    for (const stale of ['debug-suggest.png', 'debug-symbols.png']) {
      const file = path.join(output, stale);
      if (fs.existsSync(file)) fs.unlinkSync(file);
    }
    const report = { screenshots: fs.readdirSync(output).filter(name => name.endsWith('.png')), vscode: executable, mode: installedExtensions ? 'installed-vsix' : 'development', assertions: ['VBAdminx status bar visible', 'AddNumbers completion widget visible', 'VBAdminx commands visible', 'undeclaredValue diagnostic visible', 'AddNumbers hover visible', 'project status JSON visible', 'symbols view lists AddNumbers', 'clicking symbol selects AddNumbers in editor'] };
    fs.writeFileSync(path.join(root, 'output', 'screenshot-check.json'), JSON.stringify(report, null, 2));
    console.log(JSON.stringify(report));
  } finally {
    if (browser) { try { const page = browser.contexts()[0]?.pages()[0]; await page?.close(); } catch {} await browser.close(); }
    child.kill();
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
