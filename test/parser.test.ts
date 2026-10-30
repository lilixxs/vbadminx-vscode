import { test } from 'node:test';
import assert from 'node:assert/strict';
import { KEYWORDS, parseModule, tokenize } from '../src/language/parser.js';
import type { ParsedModule, ParseOptions, VbaSymbol, VbaToken } from '../src/language/model.js';

const STANDARD_URI = 'file:///c%3A/project/Module1.bas';

function parse(text: string, options: ParseOptions = {}, uri = STANDARD_URI): ParsedModule {
  return parseModule(uri, text, options);
}

function named(parsed: ParsedModule, name: string, kind?: VbaSymbol['kind']): VbaSymbol[] {
  return parsed.symbols.filter((symbol) => symbol.name === name && (kind === undefined || symbol.kind === kind));
}

function one(parsed: ParsedModule, name: string, kind?: VbaSymbol['kind']): VbaSymbol {
  const matches = named(parsed, name, kind);
  assert.equal(matches.length, 1, `expected exactly one ${kind ?? 'symbol'} named ${name}`);
  return matches[0]!;
}

function tokensOn(tokens: VbaToken[], line: number): VbaToken[] {
  return tokens.filter((token) => token.line === line);
}

function codes(parsed: ParsedModule): string[] {
  return parsed.diagnostics.map((diagnostic) => String(diagnostic.code ?? ''));
}

test('tokenize: lexical categories, escaping and UTF-16 columns', () => {
  const line0 = 'Dim s As String: s = "a""b\'c" \' trailing';
  const line1 = 'Rem whole line';
  const line2 = 'x = #1/1/2020 12:00:00 PM#';
  const line3 = 'Print #1, "v"';
  const line4 = 'y = &HFF& + 1.5E+3';
  const line5 = 'z = "😀" + 1';
  const line6 = 'Remainder = 1';
  const text = [line0, line1, line2, line3, line4, line5, line6].join('\r\n');
  const tokens = tokenize(text);

  assert.deepEqual(tokensOn(tokens, 0).map((token) => token.text), [
    'Dim', 's', 'As', 'String', ':', 's', '=', '"a""b\'c"', "' trailing",
  ]);
  const stringToken = tokensOn(tokens, 0)[7]!;
  assert.equal(stringToken.kind, 'string');
  assert.equal(stringToken.start, line0.indexOf('"a'));
  assert.equal(stringToken.end, line0.indexOf("' trailing") - 1);
  const commentToken = tokensOn(tokens, 0)[8]!;
  assert.equal(commentToken.kind, 'comment');
  assert.equal(commentToken.start, line0.indexOf("' trailing"));

  const rem = tokensOn(tokens, 1);
  assert.equal(rem.length, 1);
  assert.equal(rem[0]!.kind, 'comment');
  assert.equal(rem[0]!.text, 'Rem whole line');

  const date = tokensOn(tokens, 2)[2]!;
  assert.equal(date.kind, 'number');
  assert.equal(date.text, '#1/1/2020 12:00:00 PM#');
  assert.ok(tokensOn(tokens, 2).every((token) => token.kind !== 'identifier' || token.text === 'x'));

  assert.deepEqual(tokensOn(tokens, 3).map((token) => token.kind), ['keyword', 'operator', 'number', 'operator', 'string']);

  const hex = tokensOn(tokens, 4)[2]!;
  assert.equal(hex.kind, 'number');
  assert.equal(hex.text, '&HFF&');
  assert.equal(tokensOn(tokens, 4)[4]!.text, '1.5E+3');

  const emojiString = tokensOn(tokens, 5)[2]!;
  assert.equal(emojiString.text, '"😀"');
  assert.equal(emojiString.start, 4);
  assert.equal(emojiString.end, 8);
  assert.equal(tokensOn(tokens, 5)[3]!.start, line5.indexOf('+'));
  assert.equal(tokensOn(tokens, 5)[3]!.start, 9);

  const remainder = tokensOn(tokens, 6)[0]!;
  assert.equal(remainder.kind, 'identifier');
  assert.equal(remainder.text, 'Remainder');

  assert.ok(KEYWORDS.has('property'));
  assert.ok(KEYWORDS.has('longptr'));
  assert.ok(!KEYWORDS.has('Property'));
  assert.ok(!KEYWORDS.has('vba7'));
});

test('metadata header: class/document kind, protected lines and name location', () => {
  const source = [
    'VERSION 1.0 CLASS',
    'BEGIN',
    '  MultiUse = -1  \'True',
    'END',
    'Attribute VB_Name = "CWorker"',
    'Attribute VB_PredeclaredId = True',
    'Option Explicit',
    '',
    "' Runs the worker.",
    'Public Sub Run(ByVal count As Long)',
    'End Sub',
  ].join('\n');
  const parsed = parse(source, {}, 'file:///c%3A/project/CWorker.cls');

  assert.equal(parsed.name, 'CWorker');
  assert.equal(parsed.kind, 'document');
  assert.equal(parsed.explicit, true);
  assert.deepEqual([...parsed.protectedLines].sort((a, b) => a - b), [0, 1, 2, 3, 4, 5]);

  const moduleSymbol = one(parsed, 'CWorker', 'class');
  assert.equal(moduleSymbol.signature, 'Document CWorker');
  assert.deepEqual(moduleSymbol.selectionRange, {
    start: { line: 4, character: 21 },
    end: { line: 4, character: 28 },
  });
  assert.equal(moduleSymbol.scope, undefined);

  const run = one(parsed, 'Run', 'sub');
  assert.equal(run.scope, undefined);
  assert.equal(run.container, moduleSymbol.id);
  assert.equal(run.documentation, 'Runs the worker.');
  assert.deepEqual(run.selectionRange, {
    start: { line: 9, character: 11 },
    end: { line: 9, character: 14 },
  });
  assert.equal(run.parameters?.length, 1);
  assert.deepEqual(run.parameters?.[0], { name: 'count', typeName: 'Long', optional: false, byRef: false });
});

test('form header: designer block is protected, real code is not', () => {
  const source = [
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ',
    '   Caption         =   "UserForm1"',
    '   Begin VB.Label Label1 ',
    '      Caption         =   "Label1"',
    '   End',
    'End',
    'Attribute VB_Name = "UserForm1"',
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_PredeclaredId = True',
    'Attribute VB_Exposed = False',
    'Option Explicit',
    '',
    'Private Sub UserForm_Initialize()',
    '    Caption = "Hi"',
    'End Sub',
  ].join('\n');
  const parsed = parse(source, {}, 'file:///c%3A/project/UserForm1.frm');

  assert.equal(parsed.kind, 'form');
  assert.equal(parsed.name, 'UserForm1');
  assert.equal(parsed.explicit, true);
  for (const line of [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]) {
    assert.ok(parsed.protectedLines.has(line), `line ${line} should be protected`);
  }
  for (const line of [11, 13, 14, 15]) {
    assert.ok(!parsed.protectedLines.has(line), `line ${line} should not be protected`);
  }

  const initialize = one(parsed, 'UserForm_Initialize', 'sub');
  assert.equal(initialize.eventHandler, true);
  assert.equal(initialize.scope, undefined);
  assert.ok(parsed.folds.some((fold) => fold.startLine === 0 && fold.endLine === 10 && fold.kind === 'region'));
});

test('multi-line parameters and parameter symbols', () => {
  const source = [
    'Public Function Calc( _',
    '    ByVal a As Long, _',
    '    Optional ByRef b As String = "x", _',
    '    ParamArray rest() As Variant _',
    ') As Boolean',
    'End Function',
  ].join('\n');
  const parsed = parse(source);
  const calc = one(parsed, 'Calc', 'function');

  assert.equal(calc.scope, undefined);
  assert.equal(calc.typeName, 'Boolean');
  assert.equal(calc.signature, 'Public Function Calc(ByVal a As Long, Optional ByRef b As String, ParamArray rest() As Variant) As Boolean');
  assert.deepEqual(calc.range, {
    start: { line: 0, character: 7 },
    end: { line: 5, character: 12 },
  });
  assert.deepEqual(calc.selectionRange, {
    start: { line: 0, character: 16 },
    end: { line: 0, character: 20 },
  });
  assert.deepEqual(calc.parameters, [
    { name: 'a', typeName: 'Long', optional: false, byRef: false },
    { name: 'b', typeName: 'String', optional: true, byRef: true },
    { name: 'rest', typeName: 'Variant', optional: false, byRef: true, paramArray: true, array: true },
  ]);

  const parameters = parsed.symbols.filter((symbol) => symbol.kind === 'parameter');
  assert.deepEqual(parameters.map((symbol) => symbol.name), ['a', 'b', 'rest']);
  for (const parameter of parameters) {
    assert.equal(parameter.scope, calc.id);
    assert.equal(parameter.container, calc.id);
    assert.equal(parameter.access, 'private');
  }
  const a = one(parsed, 'a', 'parameter');
  assert.equal(a.signature, 'ByVal a As Long');
  assert.deepEqual(a.selectionRange, {
    start: { line: 1, character: 10 },
    end: { line: 1, character: 11 },
  });
});

test('scope contract for module members, locals and property accessors', () => {
  const source = [
    'Private moduleVar As Long',
    '',
    'Public Sub Worker(ByVal seed As Long)',
    '    Dim working As Long',
    '    Const stepSize As Long = 2',
    'End Sub',
    '',
    'Public Property Get Value() As Long',
    'End Property',
    'Public Property Let Value(ByVal v As Long)',
    'End Property',
    'Public Property Set Value(ByRef v As Object)',
    'End Property',
  ].join('\n');
  const parsed = parse(source);
  const moduleSymbol = one(parsed, 'Module1', 'module');

  const moduleVar = one(parsed, 'moduleVar', 'variable');
  assert.equal(moduleVar.scope, undefined);
  assert.equal(moduleVar.container, moduleSymbol.id);

  const worker = one(parsed, 'Worker', 'sub');
  assert.equal(worker.scope, undefined);

  const seed = one(parsed, 'seed', 'parameter');
  assert.equal(seed.scope, worker.id);
  assert.equal(seed.container, worker.id);

  const working = one(parsed, 'working', 'variable');
  assert.equal(working.scope, worker.id);
  assert.equal(working.container, worker.id);

  const stepSize = one(parsed, 'stepSize', 'constant');
  assert.equal(stepSize.scope, worker.id);
  assert.equal(stepSize.container, worker.id);

  const properties = named(parsed, 'Value', 'property');
  assert.equal(properties.length, 3);
  assert.equal(new Set(properties.map((symbol) => symbol.id)).size, 3);
  assert.ok(properties.some((symbol) => symbol.id.endsWith('#property:get')));
  assert.ok(properties.some((symbol) => symbol.id.endsWith('#property:let')));
  assert.ok(properties.some((symbol) => symbol.id.endsWith('#property:set')));
  const getter = properties.find((symbol) => symbol.id.endsWith('#property:get'))!;
  const setter = properties.find((symbol) => symbol.id.endsWith('#property:set'))!;
  assert.equal(getter.typeName, 'Long');
  for (const property of properties) {
    assert.equal(property.scope, undefined);
    assert.equal(property.container, moduleSymbol.id);
  }
  const setParameter = parsed.symbols.find(
    (symbol) => symbol.kind === 'parameter' && symbol.name === 'v' && symbol.scope === setter.id,
  );
  assert.ok(setParameter);
  assert.equal(setParameter.container, setter.id);
  assert.equal(setParameter.typeName, 'Object');
});

test('UDT fields and enum members use container names and scope ids', () => {
  const source = [
    'Public Type Point',
    '    X As Long, Y As Long',
    '    Z(1 To 3) As Double',
    'End Type',
    '',
    'Public Enum Color',
    '    Red = 1',
    '    Green',
    'End Enum',
  ].join('\n');
  const parsed = parse(source);

  const point = one(parsed, 'Point', 'type');
  assert.equal(point.scope, undefined);
  assert.equal(point.container, one(parsed, 'Module1', 'module').id);

  const x = one(parsed, 'X', 'field');
  const y = one(parsed, 'Y', 'field');
  const z = one(parsed, 'Z', 'field');
  for (const field of [x, y, z]) {
    assert.equal(field.container, 'Point');
    assert.equal(field.scope, point.id);
    assert.equal(field.access, 'public');
  }
  assert.equal(x.typeName, 'Long');
  assert.equal(y.typeName, 'Long');
  assert.equal(z.signature, 'Z() As Double');
  assert.equal(z.typeName, 'Double');

  const color = one(parsed, 'Color', 'enum');
  assert.equal(color.scope, undefined);
  const red = one(parsed, 'Red', 'enumMember');
  const green = one(parsed, 'Green', 'enumMember');
  for (const member of [red, green]) {
    assert.equal(member.container, 'Color');
    assert.equal(member.scope, color.id);
  }
});

test('Declare PtrSafe, Event, WithEvents and event handlers', () => {
  const source = [
    'VERSION 1.0 CLASS',
    'BEGIN',
    'END',
    'Attribute VB_Name = "CWorker"',
    'Attribute VB_PredeclaredId = False',
    'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
    'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)',
    'Public Event Changed(ByVal reason As String)',
    'Private WithEvents m_app As Application',
    'Private Sub m_app_NewWorkbook(ByVal Wb As Workbook)',
    'End Sub',
    'Private Sub Worksheet_Change()',
    'End Sub',
  ].join('\n');
  const parsed = parse(source, {}, 'file:///c%3A/project/CWorker.cls');
  const moduleSymbol = one(parsed, 'CWorker', 'class');
  assert.equal(parsed.kind, 'class');

  const tick = one(parsed, 'GetTickCount', 'function');
  assert.equal(tick.external, true);
  assert.equal(tick.scope, undefined);
  assert.equal(tick.signature, 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long');
  assert.equal(tick.typeName, 'Long');

  const sleep = one(parsed, 'Sleep', 'sub');
  assert.equal(sleep.external, true);
  assert.equal(sleep.parameters?.[0]?.name, 'ms');
  assert.equal(sleep.signature, 'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)');
  const sleepMs = one(parsed, 'ms', 'parameter');
  assert.equal(sleepMs.scope, sleep.id);
  assert.equal(sleepMs.container, sleep.id);

  const changed = one(parsed, 'Changed', 'event');
  assert.equal(changed.scope, undefined);
  assert.equal(changed.container, moduleSymbol.id);
  assert.equal(changed.parameters?.[0]?.typeName, 'String');
  const reason = one(parsed, 'reason', 'parameter');
  assert.equal(reason.scope, changed.id);
  assert.equal(reason.container, changed.id);

  const app = one(parsed, 'm_app', 'variable');
  assert.equal(app.scope, undefined);
  assert.match(app.signature, /WithEvents m_app As Application/);

  assert.equal(one(parsed, 'm_app_NewWorkbook', 'sub').eventHandler, true);
  assert.notEqual(one(parsed, 'Worksheet_Change', 'sub').eventHandler, true);

  const document = parse(source.replace('Attribute VB_PredeclaredId = False', 'Attribute VB_PredeclaredId = True'));
  assert.equal(document.kind, 'document');
  assert.equal(one(document, 'Worksheet_Change', 'sub').eventHandler, true);
});

test('conditional compilation: active branches, inactive lines and unknown constants', () => {
  const source = [
    '#Const FEATURE = False',
    '#If FEATURE Then',
    'Public Sub Hidden()',
    'End Sub',
    '#ElseIf Win64 Then',
    'Public Sub On64()',
    'End Sub',
    '#Else',
    'Public Sub On32()',
    'End Sub',
    '#End If',
    '#If UNKNOWN_FLAG Then',
    'Public Sub Mystery()',
    'End Sub',
    '#End If',
  ].join('\n');
  const parsed = parse(source, { targetBits: 64 });

  assert.deepEqual(parsed.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['On64']);
  assert.deepEqual([...parsed.inactiveLines].sort((a, b) => a - b), [2, 3, 8, 9, 12, 13]);
  assert.equal(parsed.unknownConditional, true);
  const unknownDiagnostic = parsed.diagnostics.find((diagnostic) => diagnostic.code === 'cond.unknown');
  assert.ok(unknownDiagnostic);
  assert.equal(unknownDiagnostic.severity, 3);
  assert.equal(unknownDiagnostic.range.start.line, 11);
  assert.ok(parsed.tokens.some((token) => token.text === 'Mystery'));

  const bits32 = parse(source, { targetBits: 32 });
  assert.deepEqual(bits32.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['On32']);
  assert.equal(bits32.unknownConditional, true);

  const enabled = parse([
    '#If FEATURE Then',
    'Public Sub Hidden()',
    'End Sub',
    '#ElseIf Win64 Then',
    'Public Sub On64()',
    'End Sub',
    '#End If',
  ].join('\n'), { constants: { FEATURE: true } });
  assert.deepEqual(enabled.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['Hidden']);
  assert.deepEqual([...enabled.inactiveLines], [4, 5]);
  assert.equal(enabled.unknownConditional, false);

  const unknownBoth = parse(['#If UNKNOWN_FLAG Then', 'Public Sub A()', 'End Sub', '#Else', 'Public Sub B()', 'End Sub', '#End If'].join('\n'));
  assert.equal(unknownBoth.symbols.filter((symbol) => symbol.kind === 'sub').length, 0);
  assert.deepEqual([...unknownBoth.inactiveLines].sort((a, b) => a - b), [1, 2, 4, 5]);

  const defined = parse([
    '#If defined(LEGACY) Then',
    'Public Sub Legacy()',
    'End Sub',
    '#End If',
    '#If defined(NOPE) Then',
    'Public Sub Missing()',
    'End Sub',
    '#End If',
  ].join('\n'), { constants: { LEGACY: true } });
  assert.deepEqual(defined.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['Legacy']);

  const mac = parse([
    '#If Mac Then',
    'Public Sub MacOnly()',
    'End Sub',
    '#End If',
  ].join('\n'), { constants: { Mac: true } });
  assert.deepEqual(mac.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['MacOnly']);
});

test('conditional compilation: nested branches and Option Explicit exclusion', () => {
  const source = [
    '#If Win64 Then',
    '#If Win32 Then',
    'Public Sub A()',
    'End Sub',
    '#Else',
    'Public Sub B()',
    'End Sub',
    '#End If',
    '#Else',
    'Option Explicit',
    '#End If',
    'Public Sub After()',
    'End Sub',
  ].join('\n');
  const parsed = parse(source, { targetBits: 64 });

  assert.deepEqual(parsed.symbols.filter((symbol) => symbol.kind === 'sub').map((symbol) => symbol.name), ['B', 'After']);
  assert.equal(parsed.explicit, false);
  assert.deepEqual([...parsed.inactiveLines].sort((a, b) => a - b), [2, 3, 9]);

  const unclosed = parse('#If True Then\nPublic Sub A()\nEnd Sub');
  assert.ok(codes(unclosed).includes('cond.unclosedIf'));

  const unmatched = parse('#Else\n#End If\n#ElseIf True Then');
  assert.ok(codes(unmatched).includes('cond.unmatchedElse'));
  assert.ok(codes(unmatched).includes('cond.unmatchedEndIf'));
  assert.ok(codes(unmatched).includes('cond.unmatchedElseIf'));
});

test('syntax diagnostics: missing and mismatched End, single-line If', () => {
  const missingSub = parse('Public Sub A()\n    x = 1');
  const missingSubDiagnostic = missingSub.diagnostics.find((diagnostic) => diagnostic.code === 'syntax.missingEnd');
  assert.ok(missingSubDiagnostic);
  assert.match(missingSubDiagnostic.message, /End Sub/);
  assert.deepEqual(one(missingSub, 'A', 'sub').range.end, { line: 1, character: 9 });

  const missingIf = parse('Public Sub A()\n    If True Then\n        x = 1\nEnd Sub');
  assert.deepEqual(codes(missingIf), ['syntax.missingEnd']);
  assert.match(missingIf.diagnostics[0]!.message, /End If/);
  assert.equal(missingIf.diagnostics[0]!.range.start.line, 1);

  const singleLine = parse([
    'Public Sub B()',
    '    If x Then y = 1',
    '    If x Then y = 2 Else y = 3',
    '    If x Then y = 4: z = 5',
    'End Sub',
  ].join('\n'));
  assert.deepEqual(singleLine.diagnostics, []);
  assert.deepEqual(singleLine.folds.map((fold) => [fold.startLine, fold.endLine]), [[0, 3]]);

  const unexpected = parse('End Sub\nEnd If\nElse\nEnd With');
  assert.equal(unexpected.diagnostics.filter((diagnostic) => diagnostic.code === 'syntax.unexpectedEnd').length, 3);
  assert.ok(codes(unexpected).includes('syntax.unmatchedElse'));

  const missingThen = parse('Public Sub C()\n    If True\nEnd Sub');
  assert.ok(codes(missingThen).includes('syntax.missingThen'));
});

test('Function keeps keyword-named parameter and array return type', () => {
  const source = [
    'Public Function ReadParameter(line As String) As String()',
    'End Function',
  ].join('\n');
  const parsed = parse(source);
  assert.deepEqual(parsed.diagnostics, []);

  const fn = one(parsed, 'ReadParameter', 'function');
  assert.equal(fn.parameters?.length, 1);
  assert.deepEqual(fn.parameters?.[0], { name: 'line', typeName: 'String', optional: false, byRef: true });
  assert.equal(fn.typeName, 'String()');
  assert.equal(fn.signature, 'Public Function ReadParameter(ByRef line As String) As String()');

  const parameter = one(parsed, 'line', 'parameter');
  assert.equal(parameter.typeName, 'String');
  assert.equal(parameter.scope, fn.id);
  assert.equal(parameter.container, fn.id);
});

test('single-line If with colon-separated Then statements is not a block', () => {
  const source = [
    'Public Function Guard() As Boolean',
    '    Dim a As Object',
    'Cleanup:',
    '    If Not a Is Nothing Then: a.Close: Set a = Nothing',
    '        If Not a Is Nothing Then: Set a = Nothing',
    '            Guard = True',
    '        End Function',
  ].join('\n');
  const parsed = parse(source);
  assert.deepEqual(parsed.diagnostics, []);
  const guard = one(parsed, 'Guard', 'function');
  assert.equal(guard.range.end.line, 6);
  assert.equal(guard.range.end.character, 20);
  assert.equal(one(parsed, 'a', 'variable').scope, guard.id);
  assert.deepEqual(parsed.folds.map((fold) => [fold.startLine, fold.endLine]), [[0, 5]]);
});

test('lexical diagnostics are suppressed for inactive and protected lines', () => {
  const broken = parse('Public Sub A()\n    s = "unterminated\nEnd Sub');
  const diagnostic = broken.diagnostics.find((entry) => entry.code === 'lex.unterminatedString');
  assert.ok(diagnostic);
  assert.equal(diagnostic.range.start.line, 1);

  const inactive = parse([
    '#If False Then',
    's = "unterminated',
    '#End If',
    'Public Sub A()',
    'End Sub',
  ].join('\n'));
  assert.ok(!codes(inactive).includes('lex.unterminatedString'));

  const header = parse([
    'VERSION 1.0 CLASS',
    'Attribute VB_Name = "Broken "quote"',
    'END',
  ].join('\n'));
  assert.ok(!codes(header).includes('lex.unterminatedString'));
});

test('procedure documentation from comments and Attribute descriptions', () => {
  const commentDoc = parse([
    "' First line.",
    "' Second line.",
    'Public Sub Documented()',
    'End Sub',
  ].join('\n'));
  assert.equal(one(commentDoc, 'Documented', 'sub').documentation, 'First line.\nSecond line.');

  const attributeDoc = parse([
    'Public Sub X()',
    'Attribute X.VB_Description = "Does X"',
    'End Sub',
  ].join('\n'));
  const x = one(attributeDoc, 'X', 'sub');
  assert.equal(x.documentation, 'Does X');
  assert.ok(attributeDoc.protectedLines.has(1));
  assert.equal(attributeDoc.symbols.filter((symbol) => symbol.name.includes('VB_Description')).length, 0);

  const blankSeparated = parse([
    "' Not attached.",
    '',
    'Public Sub Y()',
    'End Sub',
  ].join('\n'));
  assert.equal(one(blankSeparated, 'Y', 'sub').documentation, undefined);
});

test('module kind and name precedence', () => {
  assert.equal(parse('Option Explicit', {}, 'file:///c%3A/p/Plain.bas').kind, 'standard');
  assert.equal(parse('Option Explicit', {}, 'file:///c%3A/p/Plain.cls').kind, 'class');
  assert.equal(parse('Option Explicit', {}, 'file:///c%3A/p/Plain.frm').kind, 'form');

  const document = parse('Option Explicit', {}, 'file:///c%3A/%E5%B7%A5%E7%A8%8B/%E8%A1%A8%E6%A0%BC%E5%AF%B9%E8%B1%A1/Sheet1.cls');
  assert.equal(document.kind, 'document');
  assert.equal(document.name, 'Sheet1');

  const formDir = parse('Option Explicit', {}, 'file:///c%3A/p/%E7%AA%97%E4%BD%93/UserForm1.frm');
  assert.equal(formDir.kind, 'form');

  const header = parse([
    'Attribute VB_Name = "FromHeader"',
    'Option Explicit',
  ].join('\n'), {}, 'file:///c%3A/p/Whatever.cls');
  assert.equal(header.name, 'FromHeader');
  assert.equal(header.kind, 'class');

  const overridden = parse([
    'Attribute VB_Name = "FromHeader"',
    'Option Explicit',
  ].join('\n'), { moduleName: 'FromConfig', moduleKind: 'standard' }, 'file:///c%3A/p/Whatever.cls');
  assert.equal(overridden.name, 'FromConfig');
  assert.equal(overridden.kind, 'standard');
});

test('multi declarations, labels, colon statements and type suffixes', () => {
  const source = [
    'Private a As Long, b As String',
    'Public Const Limit = 10, Label = "x"',
    '',
    'Public Sub S()',
    '    Dim i As Long, j, k() As Integer: i = 1: j = 2',
    '    Dim s$',
    '    ReDim k(1 To 3)',
    'MyLabel: i = 3',
    '10  j = 4',
    'End Sub',
  ].join('\n');
  const parsed = parse(source);

  assert.deepEqual(named(parsed, 'a', 'variable').map((symbol) => symbol.typeName), ['Long']);
  assert.deepEqual(named(parsed, 'b', 'variable').map((symbol) => symbol.typeName), ['String']);
  const limit = one(parsed, 'Limit', 'constant');
  assert.equal(limit.typeName, 'Long');
  assert.equal(limit.access, 'public');
  const label = one(parsed, 'Label', 'constant');
  assert.equal(label.typeName, 'String');

  const worker = one(parsed, 'S', 'sub');
  assert.equal(worker.scope, undefined);
  const i = one(parsed, 'i', 'variable');
  const j = one(parsed, 'j', 'variable');
  const k = one(parsed, 'k', 'variable');
  const s = one(parsed, 's', 'variable');
  assert.equal(i.typeName, 'Long');
  assert.equal(j.typeName, undefined);
  assert.equal(k.typeName, 'Integer');
  assert.equal(s.typeName, 'String');
  for (const local of [i, j, k, s]) {
    assert.equal(local.scope, worker.id);
  }
  assert.equal(parsed.symbols.filter((symbol) => symbol.name === 'MyLabel').length, 0);
  assert.equal(parsed.symbols.filter((symbol) => symbol.name === '10').length, 0);
  assert.deepEqual(parsed.diagnostics, []);
});

test('folds: nested blocks and comment runs', () => {
  const source = [
    "' one",
    "' two",
    'Public Sub Work()',
    '    With obj',
    '        For i = 1 To 10',
    '            If i > 5 Then',
    '                Do While i < 20',
    '                    i = i + 1',
    '                Loop',
    '            End If',
    '        Next i',
    '    End With',
    'End Sub',
  ].join('\n');
  const parsed = parse(source);
  const regionFolds = parsed.folds.filter((fold) => fold.kind === 'region').map((fold) => [fold.startLine, fold.endLine]);
  assert.deepEqual(regionFolds, [
    [2, 11],
    [3, 10],
    [4, 9],
    [5, 8],
    [6, 7],
  ]);
  const commentFold = parsed.folds.find((fold) => fold.kind === 'comment');
  assert.ok(commentFold);
  assert.deepEqual([commentFold.startLine, commentFold.endLine], [0, 1]);
  assert.deepEqual(parsed.diagnostics, []);
});

test('module symbol id is stable and module text is preserved', () => {
  const source = 'Option Explicit\r\n\r\nPublic Sub Hi()\r\nEnd Sub\r\n';
  const parsed = parse(source);
  assert.equal(parsed.text, source);
  assert.equal(parsed.lines.length, 5);
  assert.equal(parsed.uri, STANDARD_URI);
  assert.equal(one(parsed, 'Module1', 'module').id, `${STANDARD_URI}#Module1#module`);
  assert.deepEqual([...parsed.inactiveLines], []);
});
