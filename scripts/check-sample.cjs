const { ProjectStore } = require('../dist/src/language/project');
const { VbaService } = require('../dist/src/language/service');
const fs = require('node:fs');
const { performance } = require('node:perf_hooks');

async function main() {
  const root = process.argv[2];
  if (!root || !fs.statSync(root).isDirectory()) throw new Error('指定一个可读的导出目录。');
  const mark = performance.now();
  const store = new ProjectStore([root]);
  await store.initialize();
  if (!store.modules.size) throw new Error('未发现工程，不能将空分析记为通过。');
  const groups = {};
  let errors = 0;
  const unknownNames = {};
  const structural = [];
  for (const module of store.modules.values()) {
    for (const diagnostic of [...module.diagnostics, ...new VbaService(store.context(module.uri)).diagnostics(module)]) {
      const code = String(diagnostic.code);
      groups[code] = (groups[code] || 0) + 1;
      if (diagnostic.severity === 1) errors++;
      if (diagnostic.code === 'undeclared') {
        const name = diagnostic.data?.name;
        if (name) unknownNames[name] = (unknownNames[name] || 0) + 1;
      }
      if (String(diagnostic.code).startsWith('syntax.') || diagnostic.code === 'argument-count') structural.push({ file: require('node:path').basename(require('node:url').fileURLToPath(module.uri)), line: diagnostic.range.start.line + 1, message: diagnostic.message });
    }
  }
  const result = { readOnly: true, files: store.modules.size, projects: store.projects.length, invalidUtf8: store.invalidFiles.size, diagnosticCounts: groups, errorDiagnostics: errors, elapsedMs: +(performance.now() - mark).toFixed(2), note: 'Diagnostics are not compiler verdicts. No Office or external type indexes were loaded. No source text is included.' };
  fs.mkdirSync('output', { recursive: true });
  fs.writeFileSync('output/sample-check.json', JSON.stringify(result, null, 2));
  console.log(JSON.stringify(result, null, 2));
  if (process.argv.includes('--debug')) console.log(JSON.stringify({ topUnknownNames: Object.entries(unknownNames).sort((a, b) => b[1] - a[1]).slice(0, 30), structural }, null, 2));
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
