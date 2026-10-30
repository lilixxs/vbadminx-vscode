const { performance } = require('node:perf_hooks');
const fs = require('node:fs');
const os = require('node:os');
const { parseModule } = require('../dist/src/language/parser');
const { VbaService } = require('../dist/src/language/service');

const started = performance.now();
const modules = [];
for (let index = 0; index < 1000; index++) {
  const lines = ['Option Explicit', `Public Function Compute${index}(ByVal inputValue As Long) As Long`, 'Dim result As Long', 'result = inputValue'];
  for (let line = 0; line < 93; line++) lines.push('result = result + 1');
  lines.push(`Compute${index} = result`, 'End Function', '');
  modules.push(parseModule(`file:///benchmark/Module${index}.bas`, lines.join('\n')));
}
const indexedMs = performance.now() - started;
const service = new VbaService({ modules, typeIndexes: [] });
const completionTimes = [];
for (let round = 0; round < 100; round++) {
  const mark = performance.now();
  service.completions(modules[round * 10], { line: 4, character: 3 });
  completionTimes.push(performance.now() - mark);
}
const mark = performance.now();
const updated = parseModule(modules[0].uri, modules[0].text.replace('result = inputValue', 'result = inputValue + 1'));
new VbaService({ modules: [updated, ...modules.slice(1)], typeIndexes: [] }).diagnostics(updated);
const incrementalMs = performance.now() - mark;
completionTimes.sort((a, b) => a - b);
const result = {
  measured: true, scope: 'in-process parser and language-service CPU; excludes extension-host startup and disk scan',
  environment: { platform: process.platform, node: process.version, cpu: os.cpus()[0]?.model, cores: os.cpus().length, memoryGB: Math.round(os.totalmem() / 1024 ** 3) },
  files: modules.length, lines: modules.reduce((sum, module) => sum + module.lines.length, 0),
  indexingMs: +indexedMs.toFixed(2), incrementalDiagnosticsMs: +incrementalMs.toFixed(2), completionP95Ms: +completionTimes[94].toFixed(2),
  caveat: 'This machine is not the 4-core/16-GB reference environment. These values do not prove the full cold-start acceptance target.',
  targets: { indexingMs: 10000, incrementalDiagnosticsMs: 500, completionP95Ms: 200 },
};
fs.mkdirSync('output', { recursive: true });
fs.writeFileSync('output/benchmark.json', JSON.stringify(result, null, 2) + '\n');
console.log(JSON.stringify(result, null, 2));
if (indexedMs > 10000 || incrementalMs > 500 || completionTimes[94] > 200) process.exitCode = 1;
