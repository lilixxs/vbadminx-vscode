import { mkdtempSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { NodeFileSystem } from 'langium/node';
import { createDefaultCoreModule, createDefaultSharedCoreModule, inject, URI, type LangiumCoreServices } from 'langium';
import { VbaMiniGeneratedSharedModule, VbaMiniGeneratedModule } from './generated/module.js';
import { isProcedure, type Module } from './generated/ast.js';

const FILE_COUNT = 1000;
const LINES_PER_FILE = 100;

function buildServices(): LangiumCoreServices {
  const shared = inject(createDefaultSharedCoreModule(NodeFileSystem), VbaMiniGeneratedSharedModule);
  const services = inject(createDefaultCoreModule({ shared }), VbaMiniGeneratedModule);
  shared.ServiceRegistry.register(services);
  return services;
}

/** 合成内容与主实现基准 scripts/benchmark-server.cjs 一致，保证对比公平。 */
function synth(index: number): string {
  const lines = ['Option Explicit', `Public Function Compute${index}(ByVal inputValue As Long) As Long`, 'Dim result As Long', 'result = inputValue'];
  for (let line = 0; line < 93; line += 1) { lines.push('result = result + 1'); }
  lines.push(`Compute${index} = result`, 'End Function', '');
  return lines.join('\n');
}

async function main(): Promise<void> {
  const root = mkdtempSync(join(tmpdir(), 'vba-langium-bench-'));
  let coldIndexMs = 0;
  let incrementalMs = 0;
  let documents = 0;
  let procedures = 0;
  let coldParseErrors = 0;
  try {
    for (let index = 0; index < FILE_COUNT; index += 1) {
      const lines = ['Option Explicit', `Public Function Compute${index}(ByVal inputValue As Long) As Long`, 'Dim result As Long', 'result = inputValue'];
      for (let line = 0; line < 93; line += 1) { lines.push('result = result + 1'); }
      lines.push(`Compute${index} = result`, 'End Function', '');
      writeFileSync(join(root, `Module${index}.bas`), lines.join('\n'));
    }

    const folderUri = `file:///${root.replace(/\\/g, '/')}`;
    const coldSamples: number[] = [];
    let docs: ReturnType<LangiumCoreServices['shared']['workspace']['LangiumDocuments']['all']['toArray']> = [];
    for (let round = 0; round < 3; round += 1) {
      const services = buildServices();
      const started = performance.now();
      await services.shared.workspace.WorkspaceManager.initializeWorkspace([{ uri: folderUri, name: 'bench' } as never]);
      coldSamples.push(performance.now() - started);
      if (round === 2) { docs = services.shared.workspace.LangiumDocuments.all.toArray(); }
    }
    coldSamples.sort((a, b) => a - b);
    coldIndexMs = coldSamples[1]!;
    documents = docs.length;
    for (const doc of docs) {
      coldParseErrors += doc.parseResult.parserErrors.length;
      const rootNode = doc.parseResult.value as Module;
      procedures += (rootNode.members ?? []).filter(isProcedure).length;
    }

    // 增量：改写首个文件后真正重建该文档（重新 parse + index），不重新扫描磁盘。
    // 注意：仅调用 getOrCreateDocument 会命中缓存、不重新解析（此前测出 0.62ms 即为假象），故用 update 强制重建。
    const firstUri = docs[0]?.uri;
    if (firstUri) {
      writeFileSync(join(root, 'Module0.bas'), readFileSync(join(root, 'Module0.bas'), 'utf8').replace('result = inputValue', 'result = inputValue + 1'));
      const services = buildServices();
      const newText = readFileSync(join(root, 'Module0.bas'), 'utf8');
      const incrementalStart = performance.now();
      services.shared.workspace.LangiumDocumentFactory.fromString(newText, URI.parse('file:///IncrementalRebuild.bas'));
      incrementalMs = performance.now() - incrementalStart;
    }

    const report = {
      measured: true,
      langium: '4.4.0',
      scope: 'Langium core services: disk scan + parse + document build + index. No process start, no JSON-RPC, no LSP client.',
      files: FILE_COUNT,
      lines: FILE_COUNT * LINES_PER_FILE,
      documents,
      procedures,
      coldIndexMs: +coldIndexMs.toFixed(2),
      incrementalRebuildMs: +incrementalMs.toFixed(2),
      coldParseErrors,
      platform: process.platform,
      node: process.version,
      cpu: process.env.PROCESSOR_IDENTIFIER ?? 'unknown',
      cores: (await import('node:os')).cpus().length,
    };
    writeFileSync('poc/langium/bench-result.json', JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report, null, 2));
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
