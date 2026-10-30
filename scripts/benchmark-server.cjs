const { spawn } = require('node:child_process');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { performance } = require('node:perf_hooks');
const { createMessageConnection, StreamMessageReader, StreamMessageWriter } = require('vscode-jsonrpc/node');

async function main() {
  const root = await fs.mkdtemp(path.join(process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Temp', 'opencode') : os.tmpdir(), 'vba-benchmark-'));
  let server;
  let connection;
  try {
    await fs.writeFile(path.join(root, 'vba-project.json'), JSON.stringify({ name: 'Synthetic benchmark' }));
    for (let index = 0; index < 1000; index++) {
      const lines = ['Option Explicit', `Public Function Compute${index}(ByVal inputValue As Long) As Long`, 'Dim result As Long', 'result = inputValue'];
      for (let line = 0; line < 93; line++) lines.push('result = result + 1');
      lines.push(`Compute${index} = result`, 'End Function', '');
      await fs.writeFile(path.join(root, `Module${index}.bas`), lines.join('\n'));
    }
    const start = performance.now();
    server = spawn(process.execPath, ['dist/server.bundle.js', '--stdio'], { stdio: ['pipe', 'pipe', 'pipe'] });
    server.stderr.on('data', data => process.stderr.write(data));
    server.on('exit', code => { if (code) console.error(`Language server exited: ${code}`); });
    connection = createMessageConnection(new StreamMessageReader(server.stdout), new StreamMessageWriter(server.stdin));
    connection.listen();
    await connection.sendRequest('initialize', { processId: process.pid, capabilities: {}, workspaceFolders: [{ uri: pathToFileURL(root).href, name: 'Synthetic benchmark' }] });
    connection.sendNotification('initialized', {});
    const status = await connection.sendRequest('vba/projectStatus');
    console.log('Benchmark: indexed', status.modules);
    const coldIndexMs = performance.now() - start;
    if (status.modules !== 1000) throw new Error(`Expected 1000 modules, got ${status.modules}`);
    const uri = pathToFileURL(path.join(root, 'Module0.bas')).href;
    const text = await fs.readFile(path.join(root, 'Module0.bas'), 'utf8');
    connection.sendNotification('textDocument/didOpen', { textDocument: { uri, languageId: 'vba', version: 1, text } });
    await connection.sendRequest('vba/projectStatus');
    console.log('Benchmark: opened document');
    const samples = [];
    for (let index = 0; index < 100; index++) {
      const mark = performance.now();
      await connection.sendRequest('textDocument/completion', { textDocument: { uri }, position: { line: 4, character: 3 } });
      samples.push(performance.now() - mark);
    }
    let resolveDiagnostics;
    const published = new Promise(resolve => { resolveDiagnostics = resolve; });
    connection.onNotification('textDocument/publishDiagnostics', params => { if (params.version === 2 && params.uri.toLowerCase() === uri.toLowerCase()) resolveDiagnostics(); });
    const editStarted = performance.now();
    connection.sendNotification('textDocument/didChange', { textDocument: { uri, version: 2 }, contentChanges: [{ text: text.replace('result = inputValue', 'result = inputValue + 1') }] });
    const timer = setTimeout(() => resolveDiagnostics('timeout'), 10000);
    const publication = await published;
    clearTimeout(timer);
    if (publication === 'timeout') throw new Error('Incremental diagnostic notification timed out.');
    const incrementalMs = performance.now() - editStarted;
    samples.sort((a, b) => a - b);
    const report = {
      measured: true, scope: 'Bundled LSP child process with real disk scan, JSON-RPC and diagnostic notification; excludes VS Code UI startup',
      platform: process.platform, node: process.version, cpu: os.cpus()[0]?.model, cores: os.cpus().length, memoryGB: Math.round(os.totalmem() / 1024 ** 3), files: status.modules, lines: 100000,
      coldProcessAndIndexMs: +coldIndexMs.toFixed(2), incrementalDiagnosticNotificationMs: +incrementalMs.toFixed(2), completionP95Ms: +samples[94].toFixed(2),
      referenceHardwareMatched: false,
    };
    await fs.writeFile('output/lsp-benchmark.json', JSON.stringify(report, null, 2));
    console.log(JSON.stringify(report, null, 2));
    await connection.sendRequest('shutdown');
    connection.sendNotification('exit');
    if (coldIndexMs > 10000 || incrementalMs > 500 || samples[94] > 200) process.exitCode = 1;
  } finally {
    connection?.dispose(); server?.kill();
    await fs.rm(root, { recursive: true, force: true });
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
