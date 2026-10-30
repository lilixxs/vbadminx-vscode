import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseModule } from '../src/language/parser.js';
import { VbaService } from '../src/language/service.js';
import { CompletionItemKind, DiagnosticSeverity } from 'vscode-languageserver/node';
import type { ParsedModule, PortableTypeIndex, VbaSymbol, VbaToken } from '../src/language/model.js';

const CRLF = '\r\n';

/** 行内首个出现位置（用于构造光标 character）。 */
const byChar = (line: string | undefined, text: string): number => {
  const index = (line ?? '').indexOf(text);
  assert.ok(index >= 0, `行中找不到 ${text}`);
  return index;
};

/** 按忽略大小写的文本查找 identifier token，第 occurrence 次出现。 */
const tokenOf = (mod: ParsedModule, text: string, occurrence = 0): VbaToken => {
  const matches = mod.tokens.filter(
    (item) => item.kind === 'identifier' && item.text.toLowerCase() === text.toLowerCase(),
  );
  assert.ok(
    matches.length >= occurrence + 1,
    `${mod.name} 中未找到标识符 ${text} #${occurrence}（实际 ${matches.length} 个）`,
  );
  return matches[occurrence]!;
};

// ---------------------------------------------------------------------------
// 夹具
// ---------------------------------------------------------------------------

const URI_A = 'file:///unit/modulea.bas';
const URI_B = 'file:///unit/moduleb.bas';

const SOURCE_A = [
  'Option Explicit',
  '',
  "' Compute：合成工具函数。",
  'Public Function Compute(value As Long) As Long',
  '    Compute = value * 2',
  'End Function',
].join(CRLF);

const SOURCE_B = [
  'Option Explicit',
  '',
  'Public Sub CallIt()',
  '    Dim result As Long',
  '    result = Compute(3)',
  'End Sub',
].join(CRLF);

const serviceOf = (...mods: ParsedModule[]): { service: VbaService; mods: ParsedModule[] } =>
  ({ service: new VbaService({ modules: mods, typeIndexes: [] }), mods });
type Svc = ReturnType<typeof serviceOf>['service'];

// ---------------------------------------------------------------------------
// LSP-09 / LSP-10：跨模块定义与引用（parse + service 集成）
// ---------------------------------------------------------------------------

test('LSP-09 跨模块公开过程：裸调用解析到定义模块', () => {
  const a = parseModule(URI_A, SOURCE_A);
  const b = parseModule(URI_B, SOURCE_B);
  const service: Svc = serviceOf(a, b).service;
  const call = tokenOf(b, 'Compute', 0);
  const resolved = service.resolve(b, call);
  assert.ok(resolved, '同工程另一标准模块的 Public 函数应可解析');
  assert.equal(resolved!.uri, URI_A, '解析结果应指向 ModuleA 的声明');
  assert.equal(resolved!.kind, 'function');
});

test('LSP-09 definition：裸调用跳转到定义的 selectionRange', () => {
  const a = parseModule(URI_A, SOURCE_A);
  const b = parseModule(URI_B, SOURCE_B);
  const service = serviceOf(a, b).service;
  const call = tokenOf(b, 'Compute', 0);
  const locations = service.definition(b, { line: call.line, character: call.start });
  assert.equal(locations.length, 1, '解析唯一成功时只返回一处定义');
  assert.equal(locations[0]!.uri, URI_A);
  const declared = tokenOf(a, 'Compute', 0);
  assert.deepEqual(locations[0]!.range, {
    start: { line: declared.line, character: declared.start },
    end: { line: declared.line, character: declared.end },
  });
});

test('LSP-10 references：包含声明、过程体赋值与跨文件调用', () => {
  const a = parseModule(URI_A, SOURCE_A);
  const b = parseModule(URI_B, SOURCE_B);
  const service = serviceOf(a, b).service;
  const declared = tokenOf(a, 'Compute', 0);
  const assigned = tokenOf(a, 'Compute', 1);
  const called = tokenOf(b, 'Compute', 0);
  const symbol = service.resolve(a, declared);
  assert.ok(symbol, '声明位置应解析为声明的符号本身');
  const references = service.references(symbol!);
  const points = references.map((reference) =>
    `${reference.uri}#${reference.range.start.line}:${reference.range.start.character}`,
  );
  assert.ok(points.includes(`${URI_A}#${declared.line}:${declared.start}`), '引用应包含 A 内声明');
  assert.ok(points.includes(`${URI_A}#${assigned.line}:${assigned.start}`), '引用应包含 A 内赋值');
  assert.ok(points.includes(`${URI_B}#${called.line}:${called.start}`), '引用应包含 B 内调用');
});

// ---------------------------------------------------------------------------
// 局部作用域与遮蔽
// ---------------------------------------------------------------------------

const URI_SHADOW = 'file:///unit/shadow.bas';

const SOURCE_SHADOW = [
  'Option Explicit',
  '',
  'Dim Value As Long',
  '',
  'Public Sub Use()',
  '    Dim Value As String',
  '    Value = "x"',
  'End Sub',
  '',
  'Public Sub Other()',
  '    Value = 5',
  'End Sub',
].join(CRLF);

test('可见性基础：过程内局部符号与模块级符号均可见', () => {
  const mod = parseModule(URI_SHADOW, SOURCE_SHADOW);
  const service = serviceOf(mod).service;
  const localDecl = tokenOf(mod, 'Value', 1); // Use 内 Dim Value
  const visible = service.visible(mod, { line: localDecl.line, character: localDecl.start });
  const names = visible.map((symbol) => symbol.name.toLowerCase());
  assert.ok(names.includes('value'), '同名局部与模块级符号应在可见集中');
});

test('局部遮蔽：过程内赋值目标解析为 String 局部变量', () => {
  const mod = parseModule(URI_SHADOW, SOURCE_SHADOW);
  const service = serviceOf(mod).service;
  const assignment = tokenOf(mod, 'Value', 2); // Use 内 Value = "x"
  const resolved = service.resolve(mod, assignment);
  assert.ok(resolved, '过程内赋值目标应可解析');
  assert.equal(resolved!.kind, 'variable');
  assert.equal(resolved!.typeName, 'String', '局部 String 应遮蔽模块级 Long');
});

test('遮蔽之外：另一过程内解析为模块级 Long', () => {
  const mod = parseModule(URI_SHADOW, SOURCE_SHADOW);
  const service = serviceOf(mod).service;
  const assignment = tokenOf(mod, 'Value', 3); // Other 内 Value = 5
  const resolved = service.resolve(mod, assignment);
  assert.ok(resolved, '模块级赋值目标应可解析');
  assert.equal(resolved!.typeName, 'Long');
});

test('声明位置解析回声明自身', () => {
  const mod = parseModule(URI_SHADOW, SOURCE_SHADOW);
  const service = serviceOf(mod).service;
  const declaration = tokenOf(mod, 'Value', 0); // 模块级 Dim Value
  const resolved = service.resolve(mod, declaration);
  assert.ok(resolved);
  assert.equal(resolved!.uri, URI_SHADOW);
  assert.equal(resolved!.typeName, 'Long');
});

// ---------------------------------------------------------------------------
// LIM-01 / COM-06：未知成员与未知类型不假报“确定错误”
// ---------------------------------------------------------------------------

const URI_UNKNOWN = 'file:///unit/unknown.bas';

const SOURCE_UNKNOWN = [
  'Option Explicit',
  '',
  'Dim Thing As UnknownThing',
  '',
  'Public Sub Go()',
  '    Thing.DoIt 1',
  'End Sub',
].join(CRLF);

test('未知类型：只给 missing-type Information，不给确定性 Error', () => {
  const mod = parseModule(URI_UNKNOWN, SOURCE_UNKNOWN);
  const service = serviceOf(mod).service;
  const diags = service.diagnostics(mod);
  const missing = diags.find((item) => item.code === 'missing-type');
  assert.ok(missing, '缺少 UnknownThing 类型信息应给 missing-type 提示');
  assert.equal(missing!.severity, DiagnosticSeverity.Information);
  const errors = diags.filter((item) => item.severity === DiagnosticSeverity.Error);
  assert.ok(errors.length === 0, `未知类型不得产生确定性 Error：${JSON.stringify(errors)}`);
});

test('未知成员：不解析为确定符号，且该行不产生 Error', () => {
  const mod = parseModule(URI_UNKNOWN, SOURCE_UNKNOWN);
  const service = serviceOf(mod).service;
  const member = tokenOf(mod, 'DoIt', 0);
  assert.equal(
    service.resolve(mod, member),
    undefined,
    '类型未知对象的成员不得被伪解析为确定符号',
  );
  const errors = service
    .diagnostics(mod)
    .filter((item) => item.severity === DiagnosticSeverity.Error && item.range.start.line === member.line);
  assert.ok(errors.length === 0, `未知成员行不得报确定性 Error：${JSON.stringify(errors)}`);
});

// ---------------------------------------------------------------------------
// LSP-05 / LSP-06：补全（UDT、外部索引、内建、关键字）
// ---------------------------------------------------------------------------

const URI_UDT = 'file:///unit/point.bas';

const SOURCE_UDT = [
  'Option Explicit',
  '',
  'Type Point',
  '    X As Double',
  '    Y As Double',
  'End Type',
  '',
  'Public Sub Try()',
  '    Dim p As Point',
  '    p.X = 1',
  'End Sub',
].join(CRLF);

test('LSP-06 UDT 成员补全：Point 字段 X / Y 应出现', () => {
  const mod = parseModule(URI_UDT, SOURCE_UDT);
  const service = serviceOf(mod).service;
  const line = mod.lines.findIndex((entry) => entry.includes('p.X = 1'));
  assert.ok(line >= 0, '夹具缺少成员访问行');
  const dotChar = byChar(mod.lines[line], 'p.');
  const completions = service.completions(mod, { line, character: dotChar + 2 });
  const labels = completions.map((item) => item.label.toLowerCase());
  assert.ok(labels.includes('x'), `成员补全应含字段 X：${labels.join(', ')}`);
  assert.ok(labels.includes('y'), `成员补全应含字段 Y：${labels.join(', ')}`);
});

const EXCEL_INDEX: PortableTypeIndex = {
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
      members: [
        { name: 'Cells', kind: 'property', returnType: 'Range' },
        {
          name: 'Run',
          kind: 'method',
          returnType: 'Variant',
          parameters: [{ name: 'Macro', typeName: 'Variant', optional: false, byRef: true }],
        },
      ],
    },
    {
      name: 'Workbook',
      kind: 'interface',
      members: [{ name: 'Name', kind: 'property', returnType: 'String' }],
    },
    {
      name: 'XlDirection',
      kind: 'enum',
      members: [{ name: 'xlUp', kind: 'constant' }],
    },
  ],
};

const URI_EXTERNAL = 'file:///unit/external.bas';

const SOURCE_EXTERNAL = [
  'Option Explicit',
  '',
  'Public Sub Drive(app As Application)',
  '    app.Run "Demo"',
  'End Sub',
].join(CRLF);

test('LSP-06 外部类型索引成员补全：Application. 提供 Cells / Run', () => {
  const mod = parseModule(URI_EXTERNAL, SOURCE_EXTERNAL);
  const service = new VbaService({ modules: [mod], typeIndexes: [EXCEL_INDEX] });
  const line = mod.lines.findIndex((entry) => entry.includes('app.Run'));
  const dotChar = byChar(mod.lines[line], 'app.');
  const completions = service.completions(mod, { line, character: dotChar + 4 });
  const labels = completions.map((item) => item.label.toLowerCase());
  assert.ok(labels.includes('cells'), `外部成员补全应含 Cells：${labels.join(', ')}`);
  assert.ok(labels.includes('run'), `外部成员补全应含 Run：${labels.join(', ')}`);
});

test('LSP-05 标识符补全：内建、关键字、外部类型与常量在列', () => {
  const mod = parseModule(URI_EXTERNAL, SOURCE_EXTERNAL);
  const service = new VbaService({ modules: [mod], typeIndexes: [EXCEL_INDEX] });
  const line = mod.lines.findIndex((entry) => entry.includes('app.Run'));
  const completions = service.completions(mod, { line, character: 0 });
  const labels = completions.map((item) => item.label);
  assert.ok(labels.includes('MsgBox'), '应包含内建 MsgBox');
  assert.ok(labels.includes('Application'), '外部类型 Application 应可补全');
  assert.ok(labels.includes('xlUp'), '外部枚举成员 xlUp 应可补全');
  const keyword = completions.find((item) => item.label === 'Dim');
  assert.ok(keyword, '应包含关键字 Dim');
  const msgBox = completions.find((item) => item.label === 'MsgBox');
  assert.equal(msgBox!.kind, CompletionItemKind.Function);
});

test('大小写不敏感：APP. 成员补全命中外部成员', () => {
  const mod = parseModule(URI_EXTERNAL, SOURCE_EXTERNAL.replace('app.Run', 'APP.Run'));
  const service = new VbaService({ modules: [mod], typeIndexes: [EXCEL_INDEX] });
  const line = mod.lines.findIndex((entry) => entry.includes('APP.Run'));
  const dotChar = byChar(mod.lines[line], 'APP.');
  const completions = service.completions(mod, { line, character: dotChar + 4 });
  const labels = completions.map((item) => item.label.toLowerCase());
  assert.ok(labels.includes('cells'), '大小写不敏感成员补全应命中 Cells');
});

// ---------------------------------------------------------------------------
// QF-02：大小写不敏感解析与 name-case 提示
// ---------------------------------------------------------------------------

test('QF-02：不同大小写引用解析到声明并给 name-case 提示', () => {
  const a = parseModule(URI_A, SOURCE_A);
  const b = parseModule(URI_B, SOURCE_B.replace('Compute(3)', 'COMPUTE(3)'));
  const service = serviceOf(a, b).service;
  const call = tokenOf(b, 'COMPUTE', 0);
  const resolved = service.resolve(b, call);
  assert.ok(resolved, '不同大小写的调用应解析到同一声明');
  assert.equal(resolved!.uri, URI_A);
  const diags = service.diagnostics(b);
  const hint = diags.find((item) => item.code === 'name-case');
  assert.ok(hint, `与声明大小写不一致应给 name-case 提示：${JSON.stringify(diags.map((item) => item.code))}`);
  assert.equal(hint!.severity, DiagnosticSeverity.Hint);
});

// ---------------------------------------------------------------------------
// REN-01 / REN-02：重命名与推断候选
// ---------------------------------------------------------------------------

const URI_REN_A = 'file:///unit/rena.bas';
const URI_REN_B = 'file:///unit/renb.bas';

const SOURCE_REN_A = [
  'Option Explicit',
  '',
  "Public Function compute(x As Long) As Long  ' compute 备注里的名字",
  '    compute = x',
  'End Function',
].join(CRLF);

const SOURCE_REN_B = [
  'Option Explicit',
  '',
  'Public Sub RunIt()',
  '    Dim r As Long',
  '    r = compute(1)',
  '    Application.Run "compute"',
  'End Sub',
].join(CRLF);

const renameTarget = (a: ParsedModule, service: Svc): { symbol: VbaSymbol; pos: { line: number; character: number } } => {
  const declaration = tokenOf(a, 'compute', 0);
  const symbol = service.resolve(a, declaration);
  assert.ok(symbol, '声明应可解析');
  return { symbol, pos: { line: declaration.line, character: declaration.start } };
};

test('REN-01 rename：跨文件确定引用参与修改，字符串与注释不触碰', () => {
  const a = parseModule(URI_REN_A, SOURCE_REN_A);
  const b = parseModule(URI_REN_B, SOURCE_REN_B);
  const service = serviceOf(a, b).service;
  const { pos } = renameTarget(a, service);
  const changes = service.rename(a, pos, 'calc');
  const uris = Object.keys(changes);
  assert.ok(uris.includes(URI_REN_A), '声明文件应参与重命名');
  assert.ok(uris.includes(URI_REN_B), 'B 内确定引用应参与重命名');
  const stringToken = b.tokens.find((item) => item.kind === 'string');
  const commentToken = a.tokens.find((item) => item.kind === 'comment');
  assert.ok(stringToken, '夹具应包含字符串字面量');
  assert.ok(commentToken, '夹具应包含注释');
  for (const [uri, edits] of Object.entries(changes)) {
    const mod = uri === URI_REN_A ? a : b;
    for (const edit of edits) {
      const matched = mod.tokens.find(
        (item) => item.line === edit.range.start.line && item.start === edit.range.start.character,
      );
      assert.ok(matched && matched.kind === 'identifier', '重命名编辑应恰好落在 identifier token 上');
      assert.equal(edit.newText, 'calc');
    }
    const insideString = (uri === URI_REN_A ? a : b).tokens.filter((item) => ['string', 'comment'].includes(item.kind)).some(
      (item) =>
        edits.some(
          (edit) =>
            edit.range.start.line === item.line &&
            edit.range.start.character >= item.start &&
            edit.range.end.character <= item.end,
        ),
    );
    assert.ok(!insideString, '重命名编辑不得触及字符串或注释');
  }
});

test('REN-02 renameCandidates：Run 字符串列为推断候选并说明理由', () => {
  const a = parseModule(URI_REN_A, SOURCE_REN_A);
  const b = parseModule(URI_REN_B, SOURCE_REN_B);
  const service = serviceOf(a, b).service;
  const { symbol } = renameTarget(a, service);
  const candidates = service.renameCandidates(symbol, 'calc');
  const runCandidate = candidates.find((item) => item.uri === URI_REN_B && /Run/i.test(item.reason));
  assert.ok(runCandidate, `Run 字符串应为推断候选：${JSON.stringify(candidates)}`);
  assert.ok(runCandidate!.reason.length > 0, '候选必须给出推断理由');
  assert.equal(runCandidate!.newText, '"calc"', '字符串候选按整体字面量给出替换');
});

test('REN-02 成员调用无法完全解析时列为推断候选', () => {
  const text = [
    'Option Explicit',
    '',
    'Public Sub Poke(p As Point)',
    '    p.compute = 1',
    'End Sub',
  ].join(CRLF);
  const mod = parseModule('file:///unit/member.bas', text);
  const service = serviceOf(mod).service;
  const declaration = tokenOf(mod, 'compute', 0);
  // 集成路径：先在另一文件声明同名符号，构造“此处调用属于另一符号”的确定引用缺失。
  const realA = parseModule(URI_A, SOURCE_A);
  const realService = serviceOf(mod, realA).service;
  const realDecl = tokenOf(realA, 'Compute', 0);
  const symbol = realService.resolve(realA, realDecl);
  assert.ok(symbol, '真实声明应可解析');
  const candidates = realService.renameCandidates(symbol, 'Renamed');
  const memberCandidate = candidates.find(
    (item) => item.uri === mod.uri && item.range.start.line === declaration.line,
  );
  assert.ok(memberCandidate, '无法完全解析的成员调用应列为推断候选');
});

test('REN-01 rename：新名为关键字时拒绝', () => {
  const mod = parseModule(URI_REN_A, SOURCE_REN_A);
  const service = serviceOf(mod).service;
  const declaration = tokenOf(mod, 'compute', 0);
  assert.throws(
    () => service.rename(mod, { line: declaration.line, character: declaration.start }, 'Sub'),
    '关键字不得用作新名称',
  );
});

test('REN-01 rename：与模块级已有符号同名冲突时拒绝', () => {
  const text = [
    'Option Explicit',
    '',
    'Dim calc As Long',
    '',
    'Public Function compute(x As Long) As Long',
    '    compute = x',
    'End Function',
  ].join(CRLF);
  const mod = parseModule('file:///unit/conflict.bas', text);
  const service = serviceOf(mod).service;
  const declaration = tokenOf(mod, 'compute', 0);
  assert.throws(
    () => service.rename(mod, { line: declaration.line, character: declaration.start }, 'Calc'),
    '与已有模块级符号大小写不敏感冲突，应拒绝重命名',
  );
});

// ---------------------------------------------------------------------------
// LIM-03 / LSP-04：Option Explicit 与参数诊断
// ---------------------------------------------------------------------------

const URI_EXPLICIT = 'file:///unit/explicit.bas';

const SOURCE_EXPLICIT = [
  'Option Explicit',
  '',
  'Public Sub Add(a As Long, b As Long)',
  'End Sub',
  '',
  'Public Sub RunAll()',
  '    missing = 1',
  '    Dim n As Long',
  '    n = "abc"',
  '    Add(1)',
  'End Sub',
].join(CRLF);

test('LIM-03：Option Explicit 下未声明符号报 undeclared Error', () => {
  const mod = parseModule(URI_EXPLICIT, SOURCE_EXPLICIT);
  assert.equal(mod.explicit, true);
  const service = serviceOf(mod).service;
  const token = tokenOf(mod, 'missing', 0);
  const undeclared = service
    .diagnostics(mod)
    .find((item) => item.code === 'undeclared' && item.range.start.line === token.line);
  assert.ok(undeclared, 'Option Explicit 下应报 undeclared');
  assert.equal(undeclared!.severity, DiagnosticSeverity.Error);
});

test('LIM-03：无 Option Explicit 时不报 undeclared 确定性错误', () => {
  const mod = parseModule(URI_EXPLICIT, SOURCE_EXPLICIT.replace('Option Explicit', 'Option Base 1'));
  assert.equal(mod.explicit, false);
  const service = serviceOf(mod).service;
  const undeclared = service.diagnostics(mod).filter((item) => item.code === 'undeclared');
  assert.ok(undeclared.length === 0, `隐式变量不得报 undeclared：${JSON.stringify(undeclared)}`);
});

test('LIM-03：已声明变量 n 不得报 undeclared', () => {
  const mod = parseModule(URI_EXPLICIT, SOURCE_EXPLICIT);
  const service = serviceOf(mod).service;
  const auth = tokenOf(mod, 'n', 1); // n = "abc" 的赋值目标
  const wrong = service
    .diagnostics(mod)
    .filter((item) => item.code === 'undeclared' && item.range.start.line === auth.line);
  assert.ok(wrong.length === 0, `过程内已声明变量不得报未声明：${JSON.stringify(wrong)}`);
});

test('LSP-04：字面量与声明类型不可判定转换时报 type-mismatch', () => {
  const mod = parseModule(URI_EXPLICIT, SOURCE_EXPLICIT);
  const service = serviceOf(mod).service;
  const mismatch = service.diagnostics(mod).filter((item) => item.code === 'type-mismatch');
  const abcLine = mod.lines.findIndex((line) => line.includes('n = "abc"'));
  assert.ok(
    mismatch.length === 1 && mismatch[0]!.range.start.line === abcLine,
    `应只有 n = "abc" 处 type-mismatch：${JSON.stringify(mismatch.map((item) => item.range))}`,
  );
});

test('LSP-04：参数数量不符报 argument-count Error', () => {
  const mod = parseModule(URI_EXPLICIT, SOURCE_EXPLICIT);
  const service = serviceOf(mod).service;
  const call = tokenOf(mod, 'Add', 1);
  const counted = service
    .diagnostics(mod)
    .find((item) => item.code === 'argument-count' && item.range.start.line === call.line);
  assert.ok(counted, `Add(1) 应报 argument-count`);
  assert.equal(counted!.severity, DiagnosticSeverity.Error);
});

test('文档模块宿主名称不假报 undeclared', () => {
  const text = [
    'Option Explicit',
    '',
    'Public Sub Paint()',
    '    Range("A1") = 1',
    'End Sub',
  ].join(CRLF);
  const mod = parseModule('file:///unit/sheet.cls', text, { moduleName: 'Sheet1', moduleKind: 'document' });
  assert.equal(mod.kind, 'document');
  const service = serviceOf(mod).service;
  const undeclared = service
    .diagnostics(mod)
    .filter((item) => item.code === 'undeclared' && item.range.start.line === 4);
  assert.ok(undeclared.length === 0, `文档宿主名称 Range 不得假报 undeclared：${JSON.stringify(undeclared)}`);
});
