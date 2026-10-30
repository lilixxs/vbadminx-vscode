const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const yauzl = require('yauzl');

const digest = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const readJson = name => JSON.parse(fs.readFileSync(path.join('output', name), 'utf8'));

function readArchive(file) {
  return new Promise((resolve, reject) => yauzl.open(file, { lazyEntries: true }, (error, zip) => {
    if (error) return reject(error);
    const entries = new Map();
    zip.on('error', reject);
    zip.on('end', () => resolve(entries));
    zip.on('entry', entry => {
      if (entry.fileName.endsWith('/')) { zip.readEntry(); return; }
      if (entry.uncompressedSize > 5 * 1024 * 1024) { zip.close(); reject(new Error('归档文件异常过大。')); return; }
      zip.openReadStream(entry, (error, stream) => {
        if (error) { reject(error); return; }
        const chunks = [];
        stream.on('data', chunk => chunks.push(chunk));
        stream.on('error', reject);
        stream.on('end', () => { entries.set(entry.fileName, Buffer.concat(chunks)); zip.readEntry(); });
      });
    });
    zip.readEntry();
  }));
}

async function main() {
  const manifest = JSON.parse(fs.readFileSync('package.json', 'utf8'));
  const file = path.join('output', `${manifest.name}-${manifest.version}.vsix`);
  const entries = await readArchive(file);
  const expected = ['package.json', 'dist/extension.bundle.js', 'dist/server.bundle.js', 'language-configuration.json', 'syntaxes/vba.tmLanguage.json', 'schemas/vba-project.schema.json', 'schemas/type-index.schema.json', 'scripts/export-com.ps1', 'GLOSSARY.md', 'docs/requirements.md', 'docs/implementation.md'];
  for (const name of expected) {
    const bytes = entries.get(`extension/${name}`);
    if (!bytes) throw new Error(`VSIX 缺少 ${name}`);
    if (digest(bytes) !== digest(fs.readFileSync(name))) throw new Error(`包内文件与当前构建不一致：${name}`);
  }
  const packagedManifest = JSON.parse(entries.get('extension/package.json').toString('utf8'));
  if (packagedManifest.version !== manifest.version) throw new Error('版本不匹配。');
  const installedDirectory = process.env.VBA_INSTALLED_EXTENSIONS || (process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Temp', 'opencode', 'vba-vsix-extensions'));
  const installedRoot = installedDirectory && path.join(installedDirectory, `${manifest.publisher}.${manifest.name}-${manifest.version}`);
  const runtimeFiles = ['dist/extension.bundle.js', 'dist/server.bundle.js', 'scripts/export-com.ps1'];
  const installedMatches = !!installedRoot && runtimeFiles.every(name => fs.existsSync(path.join(installedRoot, name)) && digest(fs.readFileSync(path.join(installedRoot, name))) === digest(entries.get(`extension/${name}`)));
  if (!installedMatches) throw new Error('隔离安装目录中的运行时与当前 VSIX 不一致。请重新安装最终包。');
  const units = readJson('unit-test-check.json');
  const api = readJson('vscode-api-check.json');
  const screenshots = readJson('screenshot-check.json');
  if (units.failed || units.exitCode || !units.tests || api.failed.length || screenshots.mode !== 'installed-vsix') throw new Error('验证结果未全部通过。');
  for (const screenshot of screenshots.screenshots) if (!fs.existsSync(path.join('output', 'screenshots', screenshot))) throw new Error(`截图不存在：${screenshot}`);
  const report = {
    generatedAt: new Date().toISOString(),
    package: { file, bytes: fs.statSync(file).size, sha256: digest(fs.readFileSync(file)), version: manifest.version, files: entries.size, requiredFilesMatchBuild: true, isolatedInstalledRuntimeMatches: installedMatches },
    checks: { unitTests: units, vscodeApi: api, screenshots, com: readJson('com-check.json'), sample: readJson('sample-check.json'), lspBenchmark: readJson('lsp-benchmark.json') },
    limitations: ['VBA 分析器采用结构解析与有限类型推断，不等同于 Office VBA 编译器。', 'macOS、Linux、Remote SSH、WSL 和容器未在实际环境验收。', '官方 Go 直连与真实 AI 聊天 UI 未验收。', 'COM record/alias 尚未表达。', '性能测试硬件与需求中的参考硬件不匹配。'],
    mainModelQuota: { status: 'unknown', reason: '主模型额度接口鉴权刷新失败，无法核实小时用量与 90% 阈值。没有用 Go 的额度代替主模型额度。' },
  };
  fs.writeFileSync('output/delivery-check.json', JSON.stringify(report, null, 2) + '\n');
  const benchmark = report.checks.lspBenchmark;
  fs.writeFileSync('output/验收报告.md', `# ${manifest.displayName ?? manifest.name} ${manifest.version} 交付验收\n\n安装包：[${path.basename(file)}](${path.basename(file)})。\n\n- 安装包大小：${report.package.bytes} 字节。\n- SHA-256：\`${report.package.sha256}\`。\n- 包内必需文件与当前构建一致。\n- 隔离安装目录中的运行时与该包一致。\n- 单元测试：${units.passed}/${units.tests} 通过。\n- VS Code ${api.vscodeVersion} API 检查：${api.passed.length} 项通过。\n- 安装版截图：${screenshots.screenshots.length} 张。\n- Windows COM 实测：${report.checks.com.types} 个类型、${report.checks.com.members} 个成员。\n- 用户样例：只读检查 ${report.checks.sample.files} 个文件，未修改源码。\n\n## 运行截图\n\n${screenshots.screenshots.map(name => `- [${name}](screenshots/${name})`).join('\n')}\n\n## 性能实测\n\n1000 文件、10 万行。冷启动及磁盘索引 ${benchmark.coldProcessAndIndexMs} 毫秒；增量诊断通知 ${benchmark.incrementalDiagnosticNotificationMs} 毫秒；补全 P95 ${benchmark.completionP95Ms} 毫秒。测试包括 LSP 子进程和 JSON-RPC，不含 VS Code UI 启动。测试机不是需求规定的参考硬件。\n\n## 验证范围\n\n${report.limitations.map(item => `- ${item}`).join('\n')}\n\n主模型小时额度：未知。${report.mainModelQuota.reason}\n\n完整证据见 [delivery-check.json](delivery-check.json)。\n`, 'utf8');
  console.log(JSON.stringify({ package: report.package, unitTests: units.tests, vscodeApiChecks: api.passed.length, screenshots: screenshots.screenshots.length }, null, 2));
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
