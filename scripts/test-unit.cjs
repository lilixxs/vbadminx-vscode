const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const files = fs.readdirSync('dist/test').filter(name => name.endsWith('.test.js')).map(name => path.join('dist/test', name));
if (!files.length) throw new Error('没有已编译测试，不能将空执行记录为通过。');
const run = spawnSync(process.execPath, ['--test', '--test-reporter=tap', ...files], { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 });
process.stdout.write(run.stdout ?? '');
process.stderr.write(run.stderr ?? '');
const count = name => Number(new RegExp(`^# ${name} (\\d+)$`, 'm').exec(run.stdout ?? '')?.[1] ?? 0);
const report = { measured: true, runner: 'node:test', node: process.version, files: files.length, tests: count('tests'), passed: count('pass'), failed: count('fail'), exitCode: run.status, generatedAt: new Date().toISOString() };
fs.mkdirSync('output', { recursive: true });
fs.writeFileSync('output/unit-test-check.json', JSON.stringify(report, null, 2) + '\n');
process.exitCode = run.status === 0 && report.tests > 0 && report.failed === 0 ? 0 : 1;
