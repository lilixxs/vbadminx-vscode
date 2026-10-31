import { test, type TestContext } from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { ProjectStore, parseProjectConfig, type ProjectConfig } from '../src/language/project.js';

const TEMP_BASE = path.join(os.tmpdir(), 'opencode');

interface Workspace {
  root: string;
  text: (rel: string, value: string) => Promise<void>;
  bytes: (rel: string, value: Uint8Array) => Promise<void>;
  object: (rel: string, value: unknown) => Promise<void>;
}

async function makeWorkspace(tag: string, t: TestContext): Promise<Workspace> {
  await fs.mkdir(TEMP_BASE, { recursive: true });
  const root = await fs.mkdtemp(path.join(TEMP_BASE, `vba-prj-${tag}-`));
  const file = (rel: string): string => path.join(root, rel);
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  return {
    root,
    text: async (rel, value) => {
      await fs.mkdir(path.dirname(file(rel)), { recursive: true });
      await fs.writeFile(file(rel), value, 'utf8');
    },
    bytes: async (rel, value) => {
      await fs.mkdir(path.dirname(file(rel)), { recursive: true });
      await fs.writeFile(file(rel), value);
    },
    object: async (rel, value) => {
      await fs.mkdir(path.dirname(file(rel)), { recursive: true });
      await fs.writeFile(file(rel), `${JSON.stringify(value)}\n`, 'utf8');
    },
  };
}

const uriOf = (root: string, rel: string): string => pathToFileURL(path.join(root, rel)).href;

const CRLF = '\r\n';

const EXCEL_INDEX = {
  formatVersion: 1,
  library: {
    name: 'Excel',
    guid: '{c05acf10-f23d-11d0-b5b2-a0c911803325}',
    major: 1,
    minor: 9,
    lcid: 0,
    bitness: 64,
  },
  types: [
    {
      name: 'Application',
      kind: 'class',
      members: [{ name: 'Cells', kind: 'property', returnType: 'Range' }],
    },
  ],
};

// ---------------------------------------------------------------------------
// PRJ-03 / CFG-01：配置文件发现、模块名/类型映射、条件编译常量与目标位数
// ---------------------------------------------------------------------------

test('PRJ-03 配置文件优先：模块名映射与 targetBits 决定条件分支', async (t) => {
  const ws = await makeWorkspace('config', t);
  await ws.object('vba-project.json', {
    name: 'ProjOne',
    targetBits: 32,
    constants: { SMOKE: true },
    modules: { '模块/Helpers.bas': { name: 'Helpers', kind: 'standard' } },
  } satisfies ProjectConfig);
  const source = [
    '#If SMOKE Then',
    'Public Sub SmokeBranch()',
    'End Sub',
    '#Else',
    'Public Sub FallbackBranch()',
    'End Sub',
    '#End If',
    '',
    '#If Win32 Then',
    'Public Sub Is32()',
    'End Sub',
    '#End If',
  ].join(CRLF);
  await ws.text('模块/Helpers.bas', source);

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  assert.equal(store.projects.length, 1, '应只识别一个工程');
  assert.equal(store.projects[0]!.root, ws.root);
  assert.equal(store.projects[0]!.name, 'ProjOne');

  const uri = uriOf(ws.root, '模块/Helpers.bas');
  const parsed = store.modules.get(uri);
  assert.ok(parsed, '扫描后模块应在索引中');
  assert.equal(parsed!.name, 'Helpers', '显式配置名优先于文件名');
  assert.equal(parsed!.kind, 'standard');
  const names = parsed!.symbols.map((symbol) => symbol.name);
  assert.ok(names.includes('SmokeBranch'), '配置常量下生效分支应被分析');
  assert.ok(!names.includes('FallbackBranch'), '未走分支应被排除');
  assert.ok(names.includes('Is32'), 'targetBits=32 时 Win32 分支应生效');
  assert.ok(parsed!.inactiveLines.size + (names.includes('Is32') ? 0 : 1) >= 1, '存在被排除的分支行');
  assert.ok(store.projects[0]!.files.has(uri), '工程文件集应记录该文件');
});

// ---------------------------------------------------------------------------
// PRJ-04 / CFG-02：无配置时的中文目录自动发现与类型推导
// ---------------------------------------------------------------------------

test('PRJ-04 四类中文目录布局自动识别；目录约定决定模块类型', async (t) => {
  const ws = await makeWorkspace('auto', t);
  await ws.text('模块/A.bas', 'Public Sub Foo()\r\nEnd Sub\r\n');
  await ws.text('类模块/B.cls', "Attribute VB_Name = \"AutoComp\"\r\nPublic Sub Foo()\r\nEnd Sub\r\n");
  await ws.text('表格对象/C.cls', 'Public Sub Foo()\r\nEnd Sub\r\n');
  await ws.text('窗体/D.frm', 'Public Sub Foo()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  assert.equal(store.projects.length, 1, `应自动识别工程：${store.projects.length}`);
  assert.equal(store.projects[0]!.name, path.basename(ws.root), '无配置时用目录名作工程名');
  const uMod = uriOf(ws.root, '模块/A.bas');
  const uCls = uriOf(ws.root, '类模块/B.cls');
  const uDoc = uriOf(ws.root, '表格对象/C.cls');
  const uFrm = uriOf(ws.root, '窗体/D.frm');
  for (const uri of [uMod, uCls, uDoc, uFrm]) {
    assert.ok(store.projectFor(uri), `文件应归属工程：${uri}`);
  }
  assert.equal(store.projectFor(uMod)!, store.projects[0]);
  assert.equal(store.modules.get(uMod)!.kind, 'standard');
  assert.equal(store.modules.get(uCls)!.kind, 'class');
  assert.equal(store.modules.get(uFrm)!.kind, 'form');
  const cls = store.modules.get(uCls)!;
  assert.equal(cls.name, 'AutoComp', '标准 VBE 头名字优先于扩展名推导');
  const doc = store.modules.get(uDoc)!;
  assert.equal(doc.kind, 'document', '表格对象目录应推导为 document');
});

// ---------------------------------------------------------------------------
// PRJ-05：无工程归属文件按单文件模式
// ---------------------------------------------------------------------------

test('PRJ-05 不在工程内的文件按单文件模式', async (t) => {
  const ws = await makeWorkspace('loose', t);
  await ws.text('模块/In.bas', 'Public Sub Inside()\r\nEnd Sub\r\n');
  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const looseUri = uriOf(ws.root, 'loose.bas');
  const parsed = store.update(looseUri, 'Option Explicit\r\nPublic Sub Loose()\r\nEnd Sub\r\n');
  assert.equal(store.projectFor(looseUri), undefined, 'loose 文件不得被归入任何工程');
  const context = store.context(looseUri);
  assert.equal(context.modules.length, 1, '单文件模式上下文只有自身');
  assert.equal(context.modules[0]!.uri, looseUri);
  assert.equal(context.typeIndexes.length, 0);
  assert.equal(store.projectDiagnostics(looseUri).length, 0);
  assert.ok(parsed.name, '以文件名推导模块名');
  void ws;
});

// ---------------------------------------------------------------------------
// PRJ-02：两工程符号隔离
// ---------------------------------------------------------------------------

test('PRJ-02 多工程：上下文只含同工程模块', async (t) => {
  const ws = await makeWorkspace('isolated', t);
  await ws.object('ProjA/vba-project.json', { name: 'ProjA' });
  await ws.object('ProjB/vba-project.json', { name: 'ProjB' });
  const marker = 'Public Sub Marker()\r\nEnd Sub\r\n';
  await ws.text('ProjA/模块/DefA.bas', marker);
  await ws.text('ProjA/模块/UseA.bas', 'Public Sub CallA()\r\n    Marker\r\nEnd Sub\r\n');
  await ws.text('ProjB/模块/DefB.bas', marker);
  await ws.text('ProjB/模块/UseB.bas', 'Public Sub CallB()\r\n    Marker\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  assert.equal(store.projects.length, 2);
  const useA = uriOf(ws.root, 'ProjA/模块/UseA.bas');
  const contextA = store.context(useA);
  assert.equal(contextA.modules.length, 2, `工程 A 上下文应只含 A 的两个模块：${contextA.modules.map((item) => item.uri)}`);
  assert.ok(
    contextA.modules.every((item) => item.uri.includes('ProjA')),
    'B 工程模块不得进入 A 的上下文',
  );

  const { VbaService } = await import('../src/language/service.js');
  const service = new VbaService(contextA);
  const target = contextA.modules.find((item) => item.uri.includes('UseA'))!;
  const call = target.tokens.find((item) => item.kind === 'identifier' && item.text.toLowerCase() === 'marker')!;
  const resolved = service.resolve(target, call);
  assert.ok(resolved, 'A 工程内 Marker 应可解析');
  assert.ok(resolved!.uri.includes('ProjA'), `解析应留在工程 A：${resolved!.uri}`);
});

// ---------------------------------------------------------------------------
// CFG-04：配置与 VBE 头冲突、重复模块名
// ---------------------------------------------------------------------------

test('CFG-04 显式模块名与 Attribute VB_Name 冲突时报告', async (t) => {
  const ws = await makeWorkspace('conflict', t);
  await ws.object('vba-project.json', {
    name: 'ProjConflict',
    modules: { '模块/Alpha.bas': { name: 'Alpha', kind: 'standard' } },
  } satisfies ProjectConfig);
  await ws.text('模块/Alpha.bas', 'Attribute VB_Name = "Beta"\r\nPublic Sub Flagship()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const uri = uriOf(ws.root, '模块/Alpha.bas');
  const parsed = store.modules.get(uri)!;
  assert.equal(parsed.name, 'Alpha', '显式配置名优先');
  const conflict = parsed.diagnostics.filter((item) => item.code === 'project-conflict');
  assert.ok(conflict.length === 1 && /冲突/.test(conflict[0]!.message), `应报名冲突：${JSON.stringify(parsed.diagnostics.map((item) => [item.code, item.message]))}`);
  assert.equal(conflict[0]!.severity, 1);
});

test('CFG-04 同工程重复模块名报告 project-conflict', async (t) => {
  const ws = await makeWorkspace('dup', t);
  await ws.object('vba-project.json', { name: 'ProjDup' } satisfies ProjectConfig);
  await ws.text('模块/DupOne.bas', 'Attribute VB_Name = "Dup"\r\nPublic Sub One()\r\nEnd Sub\r\n');
  await ws.text('模块/DupTwo.bas', 'Attribute VB_Name = "Dup"\r\nPublic Sub Two()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const one = uriOf(ws.root, '模块/DupOne.bas');
  const diags = store.projectDiagnostics(one);
  const duplicate = diags.find((item) => /模块名\s*Dup\s*重复/.test(item.message));
  assert.ok(duplicate, `重复模块名应报：${JSON.stringify(diags.map((item) => [item.code, item.message]))}`);
  assert.equal(duplicate!.code, 'project-conflict');
  assert.equal(duplicate!.severity, 1);
  assert.ok(/重名|重复/.test(duplicate!.message), '提示重命名已禁用语义');
});

// ---------------------------------------------------------------------------
// INP-03：非 UTF-8 文件停止语义分析
// ---------------------------------------------------------------------------

test('INP-03 严格 UTF-8：坏编码文件只给 encoding 提示，不给语义诊断', async (t) => {
  const ws = await makeWorkspace('encoding', t);
  await ws.object('vba-project.json', { name: 'ProjEnc' } satisfies ProjectConfig);
  // 0xC5 是 2 字节 UTF-8 序列起始，后接 0x28 非法续字节。
  await ws.bytes('模块/Bad.bas', Uint8Array.from([0x4f, 0xc5, 0x28, 0x0d, 0x0a]));

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const uri = uriOf(ws.root, '模块/Bad.bas');
  assert.ok(store.invalidFiles.has(uri), '文件应记入 invalidFiles');
  const parsed = store.modules.get(uri);
  assert.ok(parsed);
  assert.ok(
    parsed!.diagnostics.length === 1 && parsed!.diagnostics[0]!.code === 'encoding',
    `应只有 encoding 诊断：${JSON.stringify(parsed!.diagnostics.map((item) => item.code))}`,
  );
  const context = store.context(uri);
  assert.equal(context.modules.length, 0, '语义分析已暂停：坏编码文件不进入上下文');
});

test('INP-02：UTF-8 BOM 允许且解析正常', async (t) => {
  const ws = await makeWorkspace('bom', t);
  await ws.object('vba-project.json', { name: 'ProjBom' } satisfies ProjectConfig);
  await ws.bytes('模块/Bom.bas', new TextEncoder().encode(`\uFEFFPublic Sub WithBom()\r\nEnd Sub\r\n`));
  const store = new ProjectStore([ws.root]);
  await store.initialize();
  const uri = uriOf(ws.root, '模块/Bom.bas');
  const parsed = store.modules.get(uri);
  assert.ok(parsed, `工程扫描应加载 BOM 文件：${[...store.modules.keys()].join(', ')}`);
  assert.equal(store.invalidFiles.has(uri), false);
  assert.ok(parsed!.symbols.some((symbol) => symbol.name === 'WithBom'));
});

// ---------------------------------------------------------------------------
// 类型索引加载与引用一致性（CFG-04 / COM-03）
// ---------------------------------------------------------------------------

test('类型索引加载：有效入列、坏文件与缺失引用报错', async (t) => {
  const ws = await makeWorkspace('index', t);
  await ws.object('indexes/Excel.json', EXCEL_INDEX);
  await ws.text('indexes/Broken.json', '{"formatVersion": 1');
  await ws.object('vba-project.json', {
    name: 'ProjIndex',
    typeIndexes: ['indexes/Excel.json', 'indexes/Broken.json'],
    references: [
      { name: 'Excel', guid: '{c05acf10-f23d-11d0-b5b2-a0c911803325}', major: 1, minor: 9, lcid: 0, bitness: 64 },
      { name: 'Deckard', guid: '{12345678-90ab-cdef-1234-567890abcdef}', major: 2, minor: 0, lcid: 0, bitness: 64 },
    ],
  } satisfies ProjectConfig);
  await ws.text('模块/Use.bas', 'Public Sub Foo()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const project = store.projects[0]!;
  assert.equal(project.indexes.length, 1, '仅有有效索引入列');
  assert.equal(project.indexes[0]!.library.name, 'Excel');
  const messages = project.errors.join('\n');
  assert.ok(/Broken\.json/.test(messages) && /无法加载类型索引/.test(messages), `坏索引应报错：${messages}`);
  assert.ok(/类型信息缺失：Deckard 2\.0/.test(messages), `缺失引用应报错：${messages}`);
  const useUri = uriOf(ws.root, '模块/Use.bas');
  const context = store.context(useUri);
  assert.equal(context.typeIndexes.filter(index => index.library.name === 'Excel').length, 1);
  assert.ok(context.typeIndexes.some(index => index.library.name.includes('Host Model')), '应补充内建宿主模型');
});

test('CFG-04 类型索引与显式引用版本不匹配时报错且不入列', async (t) => {
  const ws = await makeWorkspace('version', t);
  await ws.object('indexes/Excel.json', EXCEL_INDEX);
  await ws.object('vba-project.json', {
    name: 'ProjVersion',
    typeIndexes: ['indexes/Excel.json'],
    references: [
      { name: 'Excel', guid: '{c05acf10-f23d-11d0-b5b2-a0c911803325}', major: 1, minor: 8, lcid: 0, bitness: 64 },
    ],
  } satisfies ProjectConfig);
  await ws.text('模块/Use.bas', 'Public Sub Foo()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  const project = store.projects[0]!;
  assert.equal(project.indexes.length, 0, '版本不匹配的索引不得入列');
  assert.ok(
    /与显式引用版本不匹配/.test(project.errors.join('\n')),
    `应报版本不匹配：${JSON.stringify(project.errors)}`,
  );
});

// ---------------------------------------------------------------------------
// 工程边界：嵌套工程互不串扫
// ---------------------------------------------------------------------------

test('嵌套工程：父扫描跳过子工程根，各自归属正确', async (t) => {
  const ws = await makeWorkspace('nested', t);
  await ws.object('ProjTop/vba-project.json', { name: 'ProjTop' });
  await ws.object('ProjTop/sub/vba-project.json', { name: 'ProjSub' });
  await ws.text('ProjTop/模块/Top.bas', 'Public Sub TopProc()\r\nEnd Sub\r\n');
  await ws.text('ProjTop/sub/模块/Low.bas', 'Public Sub LowProc()\r\nEnd Sub\r\n');

  const store = new ProjectStore([ws.root]);
  await store.initialize();

  assert.equal(store.projects.length, 2);
  const topUri = uriOf(ws.root, 'ProjTop/模块/Top.bas');
  const lowUri = uriOf(ws.root, 'ProjTop/sub/模块/Low.bas');
  const sub = store.projects.find((project) => project.root === path.join(ws.root, 'ProjTop', 'sub'))!;
  const top = store.projects.find((project) => project.root === path.join(ws.root, 'ProjTop'))!;
  assert.ok(sub && top);
  assert.ok(top.files.has(topUri));
  assert.ok(!top.files.has(lowUri), '父工程不得收编子工程文件');
  assert.ok(sub.files.has(lowUri));
  assert.equal(store.projectFor(lowUri), sub);
});

// ---------------------------------------------------------------------------
// parseProjectConfig 验证
// ---------------------------------------------------------------------------

const VALID_GUID = '{12345678-90ab-cdef-1234-567890abcdef}';

test('parseProjectConfig：非法输入拒绝，合法输入通过', () => {
  assert.throws(() => parseProjectConfig(null));
  assert.throws(() => parseProjectConfig('text'));
  assert.throws(() => parseProjectConfig(42));
  assert.throws(() => parseProjectConfig([]));
  assert.throws(() => parseProjectConfig({ targetBits: 16 }));
  assert.throws(() => parseProjectConfig({ name: 1 }));
  assert.throws(() => parseProjectConfig({ host: {} }));
  assert.throws(() => parseProjectConfig({ typeIndexes: 'a.json' }));
  assert.throws(() => parseProjectConfig({ constants: { bad: {} } }));
  assert.throws(() => parseProjectConfig({ constants: 'x' }));
  assert.throws(() => parseProjectConfig({ modules: 'x' }));
  assert.throws(() => parseProjectConfig({ modules: { a: null } }));
  assert.throws(() => parseProjectConfig({ modules: { a: { kind: 'bogus' } } }));
  assert.throws(() => parseProjectConfig({ modules: { a: { name: '1name' } } }));
  assert.throws(() => parseProjectConfig({ references: 'x' }));
  assert.throws(() => parseProjectConfig({ references: [{}] }));
  assert.throws(() => parseProjectConfig({ references: [{ name: 'L', guid: 5, major: 1, minor: 0, lcid: 0, bitness: 32 }] }));
  assert.throws(() => parseProjectConfig({ references: [{ name: 'L', guid: VALID_GUID, major: 1, minor: 0, lcid: 0, bitness: 96 }] }));

  const ok = parseProjectConfig({
    name: 'P',
    host: 'Excel',
    targetBits: 64,
    constants: { Flag: false, Count: 3, Label: '@' },
    modules: { '模块/C.bas': { name: 'C', kind: 'form' } },
    typeIndexes: ['i.json'],
    references: [{ name: 'L', guid: VALID_GUID, major: 1, minor: 0, lcid: 0, bitness: 64 }],
  });
  assert.equal(ok.targetBits, 64);
});
