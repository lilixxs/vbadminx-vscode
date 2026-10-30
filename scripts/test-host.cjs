const path = require('node:path');
const fs = require('node:fs');
const { runTests } = require('@vscode/test-electron');

async function main() {
  // 可显式覆盖，避免 PATH 中的旧版 code 被误用。
  const installed = process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Programs', 'Microsoft VS Code', 'Code.exe');
  const executable = process.env.VSCODE_EXECUTABLE_PATH || (installed && fs.existsSync(installed) ? installed : undefined);
  const profileRoot = path.resolve(__dirname, '../.vscode-test');
  fs.mkdirSync(profileRoot, { recursive: true });
  const profile = fs.mkdtempSync(path.join(profileRoot, 'profile-'));
  try { await runTests({
    vscodeExecutablePath: executable,
    version: '1.104.3',
    extensionDevelopmentPath: path.resolve(__dirname, '..'),
    extensionTestsPath: path.resolve(__dirname, '../dist/test/host/index.js'),
    launchArgs: [
      path.resolve(__dirname, '../examples'),
      '--disable-extensions',
      '--disable-workspace-trust',
      '--skip-welcome',
      '--skip-release-notes',
      '--user-data-dir', profile,
    ],
  }); } finally { fs.rmSync(profile, { recursive: true, force: true, maxRetries: 3, retryDelay: 100 }); }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
