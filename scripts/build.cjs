const { build } = require('esbuild');

async function main() {
  await build({ entryPoints: ['src/extension.ts'], outfile: 'dist/extension.bundle.js', bundle: true, platform: 'node', target: 'node20', format: 'cjs', external: ['vscode'], sourcemap: false });
  await build({ entryPoints: ['src/server.ts'], outfile: 'dist/server.bundle.js', bundle: true, platform: 'node', target: 'node20', format: 'cjs', sourcemap: false });
}
main().catch(error => { console.error(error); process.exitCode = 1; });
