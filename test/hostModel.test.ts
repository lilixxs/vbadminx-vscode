import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import test from 'node:test';
import { parseModule } from '../src/language/parser.js';
import { ProjectStore } from '../src/language/project.js';
import { VbaService } from '../src/language/service.js';
import {
  EXCEL_HOST_LIBRARY_NAME,
  excelHostGlobals,
  excelHostIndex,
  isExcelHostIndex,
  usesExcelHostModel,
} from '../src/language/hostModel.js';

const CRLF = '\r\n';

const EXCEL_SOURCE = [
  'Option Explicit',
  '',
  'Public Sub Demo()',
  '    Dim ws As Worksheet',
  '    Dim wb As Workbook',
  '    Dim lastRow As Long',
  '    Set wb = ThisWorkbook',
  '    Set ws = wb.Worksheets("Sheet1")',
  '    ws.Range("A1").Value = 1',
  '    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
  '    ws.Range("A1").Font.Bold = True',
  'End Sub',
].join(CRLF);

const lineOf = (source: string, text: string): number => {
  const index = source.split(CRLF).findIndex(line => line.includes(text));
  assert.ok(index >= 0, `找不到行：${text}`);
  return index;
};

test('内建 Excel 宿主索引包含常见对象与成员', () => {
  const index = excelHostIndex(64);
  assert.equal(index.library.name, EXCEL_HOST_LIBRARY_NAME);
  assert.equal(index.library.major, 16);
  assert.equal(index.library.bitness, 64);
  assert.ok(isExcelHostIndex(index));
  const names = new Set(index.types.map(type => type.name));
  for (const name of ['Application', 'Workbook', 'Worksheet', 'Range', 'Worksheets', 'ListObject', 'Chart', 'Font', 'Interior']) {
    assert.ok(names.has(name), `缺少类型 ${name}`);
  }
  const worksheet = index.types.find(type => type.name === 'Worksheet')!;
  const members = new Set(worksheet.members.map(member => member.name));
  for (const name of ['Range', 'Cells', 'Rows', 'Columns', 'UsedRange', 'Activate', 'Copy']) {
    assert.ok(members.has(name), `Worksheet 缺少成员 ${name}`);
  }
  assert.equal(index.types.find(type => type.name === 'Worksheet')!.members.find(member => member.name === 'Range')!.returnType, 'Range');
});

test('宿主全局符号包含 ThisWorkbook 与常用常量', () => {
  const globals = excelHostGlobals();
  const byName = new Map(globals.map(symbol => [symbol.name.toLowerCase(), symbol]));
  assert.equal(byName.get('thisworkbook')!.typeName, 'Workbook');
  assert.equal(byName.get('worksheets')!.typeName, 'Worksheet');
  assert.equal(byName.get('xlup')!.kind, 'constant');
  assert.equal(byName.get('xlup')!.typeName, 'Long');
  assert.ok(globals.every(symbol => symbol.external));
  assert.equal(usesExcelHostModel('Excel'), true);
  assert.equal(usesExcelHostModel('Microsoft Excel'), true);
  assert.equal(usesExcelHostModel(undefined), true);
  assert.equal(usesExcelHostModel('Word'), false);
});

test('Worksheet/Workbook 不再产生 missing-type 诊断，成员可解析', () => {
  const uri = 'file:///unit/excel.bas';
  const module = parseModule(uri, EXCEL_SOURCE);
  const service = new VbaService({ modules: [module], typeIndexes: [excelHostIndex(64)], hostGlobals: excelHostGlobals() });
  const diagnostics = service.diagnostics(module);
  assert.equal(diagnostics.filter(item => item.code === 'missing-type').length, 0, JSON.stringify(diagnostics));
  assert.equal(diagnostics.filter(item => item.code === 'undeclared').length, 0, JSON.stringify(diagnostics));
  assert.equal(diagnostics.filter(item => item.code === 'missing-external-symbol').length, 0, JSON.stringify(diagnostics));

  const line = lineOf(EXCEL_SOURCE, 'ws.Range("A1").Value = 1');
  const character = EXCEL_SOURCE.split(CRLF)[line]!.indexOf('ws.') + 3;
  const labels = service.completions(module, { line, character }).map(item => item.label);
  for (const expected of ['Range', 'Cells', 'Rows', 'UsedRange', 'Name']) {
    assert.ok(labels.includes(expected), `ws. 补全缺少 ${expected}：${labels.join(', ')}`);
  }

  const workbookLine = lineOf(EXCEL_SOURCE, 'Set ws = wb.Worksheets("Sheet1")');
  const workbookCharacter = EXCEL_SOURCE.split(CRLF)[workbookLine]!.indexOf('wb.') + 3;
  const workbookLabels = service.completions(module, { line: workbookLine, character: workbookCharacter }).map(item => item.label);
  for (const expected of ['Worksheets', 'Sheets', 'Name', 'FullName']) {
    assert.ok(workbookLabels.includes(expected), `wb. 补全缺少 ${expected}：${workbookLabels.join(', ')}`);
  }
});

const INDEX_FIXTURE = {
  formatVersion: 1,
  library: { name: 'Excel', guid: '{00020813-0000-0000-c000-000000000046}', major: 1, minor: 9, lcid: 0, bitness: 64 },
  types: [{ name: 'Worksheet', kind: 'class', members: [{ name: 'CustomMember', kind: 'method' }] }],
};

async function workspace(tag: string, host: string | undefined, options: { explicitIndex?: boolean } = {}): Promise<{ root: string; uri: string }> {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), `vba-host-${tag}-`));
  await fs.writeFile(path.join(root, 'vba-project.json'), JSON.stringify({
    name: tag,
    ...(host ? { host } : {}),
    targetBits: 64,
    constants: {},
    modules: {},
    typeIndexes: options.explicitIndex ? ['excel-index.json'] : [],
  }), 'utf8');
  await fs.writeFile(path.join(root, 'Module1.bas'), EXCEL_SOURCE, 'utf8');
  if (options.explicitIndex) { await fs.writeFile(path.join(root, 'excel-index.json'), JSON.stringify(INDEX_FIXTURE), 'utf8'); }
  return { root, uri: pathToFileURL(path.join(root, 'Module1.bas')).href };
}

test('工程配置 host=Excel 时自动注入宿主模型', async () => {
  const { root, uri } = await workspace('excel', 'Excel');
  try {
    const store = new ProjectStore([root]);
    await store.initialize();
    const context = store.context(uri);
    assert.ok(context.typeIndexes.some(index => index.library.name === EXCEL_HOST_LIBRARY_NAME));
    assert.ok(context.hostGlobals?.some(symbol => symbol.name === 'ThisWorkbook'));
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('非 Excel 宿主不注入宿主模型', async () => {
  const { root, uri } = await workspace('word', 'Word');
  try {
    const store = new ProjectStore([root]);
    await store.initialize();
    const context = store.context(uri);
    assert.equal(context.typeIndexes.some(index => index.library.name === EXCEL_HOST_LIBRARY_NAME), false);
    assert.equal(context.hostGlobals, undefined);
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('显式索引中的同名类型优先于内建模型', async () => {
  const { root, uri } = await workspace('explicit', 'Excel', { explicitIndex: true });
  try {
    const store = new ProjectStore([root]);
    await store.initialize();
    const context = store.context(uri);
    const providers = context.typeIndexes.filter(index => index.types.some(type => type.name === 'Worksheet'));
    assert.equal(providers.length, 1, 'Worksheet 只能由一个索引提供');
    assert.equal(providers[0]!.library.name, 'Excel');
    const host = context.typeIndexes.find(index => isExcelHostIndex(index));
    assert.ok(host);
    assert.equal(host!.types.some(type => type.name === 'Worksheet'), false);
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
});
