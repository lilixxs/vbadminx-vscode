import { test } from 'node:test';
import assert from 'node:assert/strict';
import { formatVba } from '../src/language/formatter.js';

function lines(...items: string[]): string {
  return items.join('\r\n');
}

test('默认 4 空格缩进、关键字标准大小写，并保留 BOM 与 CRLF', () => {
  const source = '\uFEFFoption explicit\r\n\r\npublic sub main()\r\nif x then\r\nx = 1\r\nend if\r\nend sub\r\n';
  const expected = '\uFEFFOption Explicit\r\n\r\nPublic Sub main()\r\n    If x Then\r\n        x = 1\r\n    End If\r\nEnd Sub\r\n';
  assert.equal(formatVba(source), expected);
});

test('保留每行原有换行形式，不把 LF 统一成 CRLF', () => {
  const source = 'Sub S()\r\nx = 1\nEnd Sub\r\n';
  assert.equal(formatVba(source), 'Sub S()\r\n    x = 1\nEnd Sub\r\n');
  assert.equal(formatVba('sub s()\nx = 1\nend sub\n'), 'Sub s()\n    x = 1\nEnd Sub\n');
});

test('字符串与注释内容逐字符保留，字符串内的关键字与引号不被改写', () => {
  const source = lines(
    'Sub S()',
    "s = \"  Keep  ; ' not a comment  if then \"  ' 尾部注释 if then  ",
    "' 只有注释 if then end  ",
    'End Sub',
  );
  const expected = lines(
    'Sub S()',
    "    s = \"  Keep  ; ' not a comment  if then \"  ' 尾部注释 if then  ",
    "    ' 只有注释 if then end  ",
    'End Sub',
  );
  assert.equal(formatVba(source), expected);
});

test('成员访问后的词不改写大小写', () => {
  const source = lines('Sub S()', 'obj.name = 1', 'obj.width = 2', 'End Sub');
  assert.equal(formatVba(source), lines('Sub S()', '    obj.name = 1', '    obj.width = 2', 'End Sub'));
});

test('Attribute 与窗体设计段默认原样保留', () => {
  const source = lines(
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain',
    '   Caption         =   " Form1 "',
    '   Begin VB.CommandButton ok ',
    '      caption   =   "ok"',
    '   End',
    'End',
    'Attribute VB_Name = "frmMain"',
    'Option Explicit',
    'Private Sub ok_Click()',
    'x = 1',
    'End Sub',
  );
  const output = formatVba(source);
  const expected = lines(
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain',
    '   Caption         =   " Form1 "',
    '   Begin VB.CommandButton ok ',
    '      caption   =   "ok"',
    '   End',
    'End',
    'Attribute VB_Name = "frmMain"',
    'Option Explicit',
    'Private Sub ok_Click()',
    '    x = 1',
    'End Sub',
  );
  assert.equal(output, expected);
});

test('formatDesign 仅调整设计段排版，不改字符串与 Attribute', () => {
  const source = lines(
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain',
    '   Caption         =   " Form1 "',
    '   Begin VB.CommandButton ok ',
    '      caption   =   "ok"',
    '   End',
    'End',
    'Attribute VB_Name = "frmMain"',
  );
  const expected = lines(
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain',
    '    Caption         =   " Form1 "',
    '    Begin VB.CommandButton ok ',
    '        caption   =   "ok"',
    '    End',
    'End',
    'Attribute VB_Name = "frmMain"',
  );
  assert.equal(formatVba(source, { formatDesign: true }), expected);
});

test('选区只格式化指定行，缩进基于全文上下文', () => {
  const source = lines('sub s()', 'if a then', 'x = 1', 'end if', 'end sub');
  const output = formatVba(source, {}, { start: 2, end: 2 });
  assert.equal(output, lines('sub s()', 'if a then', '        x = 1', 'end if', 'end sub'));

  const whole = formatVba(source);
  assert.equal(whole, lines('Sub s()', '    If a Then', '        x = 1', '    End If', 'End Sub'));

  assert.equal(formatVba(source, {}, { start: 3, end: 1 }), source);
});

test('续行按语句缩进加一级，并保留续行标志', () => {
  const source = lines('Sub S()', 'x = 1 + _', '  2', 'End Sub');
  assert.equal(formatVba(source), lines('Sub S()', '    x = 1 + _', '        2', 'End Sub'));
});

test('单行 If 与冒号分隔语句不改变后续缩进', () => {
  const source = lines('Sub S()', 'if a then b = 1: c = 2', 'x = 3', 'End Sub');
  assert.equal(formatVba(source), lines('Sub S()', '    If a Then b = 1: c = 2', '    x = 3', 'End Sub'));

  const withLabel = lines('Sub S()', 'start:', 'x = 1', 'End Sub');
  assert.equal(formatVba(withLabel), lines('Sub S()', 'start:', '    x = 1', 'End Sub'));
});

test('Select Case、Case 与 End Select 层级正确', () => {
  const source = lines(
    'Sub S()',
    'select case x',
    'case 1',
    'a = 1',
    'case else',
    'a = 2',
    'end select',
    'End Sub',
  );
  assert.equal(
    formatVba(source),
    lines(
      'Sub S()',
      '    Select Case x',
      '        Case 1',
      '            a = 1',
      '        Case Else',
      '            a = 2',
      '    End Select',
      'End Sub',
    ),
  );
});

test('If/ElseIf/Else、Do/Loop、For/Next、With、Type、Enum、Property 块均正确配对', () => {
  const source = lines(
    'Private Type TPoint',
    'x As Long',
    'End Type',
    '',
    'Public Enum EColor',
    'red = 1',
    'End Enum',
    '',
    'Property Get Name() As String',
    'Name = m_name',
    'End Property',
    '',
    'Sub S()',
    'if a then',
    'x = 1',
    'elseif b then',
    'x = 2',
    'else',
    'x = 3',
    'end if',
    'for i = 1 to 10',
    'with obj',
    '.value = i',
    'end with',
    'next i',
    'do while i > 0',
    'i = i - 1',
    'loop',
    'End Sub',
  );
  const expected = lines(
    'Private Type TPoint',
    '    x As Long',
    'End Type',
    '',
    'Public Enum EColor',
    '    red = 1',
    'End Enum',
    '',
    'Property Get Name() As String',
    '    Name = m_name',
    'End Property',
    '',
    'Sub S()',
    '    If a Then',
    '        x = 1',
    '    ElseIf b Then',
    '        x = 2',
    '    Else',
    '        x = 3',
    '    End If',
    '    For i = 1 To 10',
    '        With obj',
    '            .value = i',
    '        End With',
    '    Next i',
    '    Do While i > 0',
    '        i = i - 1',
    '    Loop',
    'End Sub',
  );
  assert.equal(formatVba(source), expected);
});

test('条件编译指令不参与块层级，且不被改写语义', () => {
  const source = lines(
    'Sub S()',
    '#if DEBUG_MODE then',
    'x = 1',
    '#else',
    'x = 2',
    '#end if',
    'End Sub',
  );
  const expected = lines(
    'Sub S()',
    '    #If DEBUG_MODE Then',
    '        x = 1',
    '    #Else',
    '        x = 2',
    '    #End If',
    'End Sub',
  );
  assert.equal(formatVba(source), expected);
});

test('keywordCase 关闭时只改缩进', () => {
  const source = lines('sub s()', 'if a then', 'x = 1', 'end if', 'end sub');
  assert.equal(
    formatVba(source, { keywordCase: false }),
    lines('sub s()', '    if a then', '        x = 1', '    end if', 'end sub'),
  );
});

test('insertSpaces=false 使用制表符，tabSize 影响空格数', () => {
  const source = lines('Sub S()', 'If a Then', 'x = 1', 'End If', 'End Sub');
  assert.equal(
    formatVba(source, { insertSpaces: false }),
    lines('Sub S()', '\tIf a Then', '\t\tx = 1', '\tEnd If', 'End Sub'),
  );
  assert.equal(
    formatVba(source, { tabSize: 2 }),
    lines('Sub S()', '  If a Then', '    x = 1', '  End If', 'End Sub'),
  );
});

test('formatVba 对空文本与仅 BOM 文本保持原样', () => {
  assert.equal(formatVba(''), '');
  assert.equal(formatVba('\uFEFF'), '\uFEFF');
});

test('未闭合字符串所在行原样保留', () => {
  const source = lines('Sub S()', '  s = "未闭合', 'End Sub');
  const output = formatVba(source);
  assert.equal(output, lines('Sub S()', '  s = "未闭合', 'End Sub'));
});

test('Rem 注释内容不被改写，冒号后同样识别', () => {
  const source = lines('Sub S()', 'rem if then end 保留原样', 'x = 1: Rem for next 保留', 'End Sub');
  assert.equal(
    formatVba(source),
    lines('Sub S()', '    rem if then end 保留原样', '    x = 1: Rem for next 保留', 'End Sub'),
  );
});

test('While/Wend 与 Do Until/Loop 配对', () => {
  const source = lines(
    'Sub S()',
    'while x > 0',
    'x = x - 1',
    'wend',
    'do until y > 0',
    'y = y + 1',
    'loop',
    'End Sub',
  );
  assert.equal(
    formatVba(source),
    lines(
      'Sub S()',
      '    While x > 0',
      '        x = x - 1',
      '    Wend',
      '    Do Until y > 0',
      '        y = y + 1',
      '    Loop',
      'End Sub',
    ),
  );
});

test('lineWrap 可在无括号调用的顶层逗号处续行，但不拆 Set 语句', () => {
  const call = `MsgBox ${Array.from({ length: 10 }, (_v, i) => `"message ${i}"`).join(', ')}`;
  const source = lines('Sub S()', call, 'End Sub');
  const output = formatVba(source, { lineWrap: true });
  assert.ok(output.includes(' _\r\n'));
  for (const line of output.split('\r\n')) assert.ok(line.length <= 120, `超长行: ${line}`);

  const setLine = `Set x = ${'Obj.'.repeat(30)}Value(1, 2)`;
  const setSource = lines('Sub S()', setLine, 'End Sub');
  assert.equal(formatVba(setSource, { lineWrap: true }), formatVba(setSource));
});

test('reorderDeclarations 仅重排安全的纯 Dim 区域', () => {
  const source = lines(
    'Sub S()',
    'Dim zebra As Long',
    'Dim apple As Long',
    'Dim mango As Long',
    'Dim y As Long: Dim w As Long',
    'Const fixedValue As Long = 3',
    'Dim offset As Long = 4',
    'Dim count As Long',
    'End Sub',
  );
  const expected = lines(
    'Sub S()',
    '    Dim apple As Long',
    '    Dim mango As Long',
    '    Dim zebra As Long',
    '    Dim y As Long: Dim w As Long',
    '    Const fixedValue As Long = 3',
    '    Dim offset As Long = 4',
    '    Dim count As Long',
    'End Sub',
  );
  assert.equal(formatVba(source, { reorderDeclarations: true }), expected);
  // 默认不重排
  assert.equal(formatVba(source).includes('Dim zebra As Long\r\n    Dim apple'), true);
});

test('reorderDeclarations 遇到重复名或 As New 时保持原样', () => {
  const source = lines(
    'Sub S()',
    'Dim b As Long',
    'Dim a As Long',
    'Dim a As Long',
    'End Sub',
  );
  assert.equal(
    formatVba(source, { reorderDeclarations: true }),
    lines('Sub S()', '    Dim b As Long', '    Dim a As Long', '    Dim a As Long', 'End Sub'),
  );

  const withNew = lines('Sub S()', 'Dim b As New Collection', 'Dim a As Long', 'End Sub');
  assert.equal(
    formatVba(withNew, { reorderDeclarations: true }),
    lines('Sub S()', '    Dim b As New Collection', '    Dim a As Long', 'End Sub'),
  );
});

test('lineWrap 只在安全逗号处续行且每行不超过 120 字符', () => {
  const args = Array.from({ length: 12 }, (_value, index) => `"value${index}"`).join(', ');
  const source = lines('Sub S()', `x = Foo(${args})`, 'End Sub');
  const output = formatVba(source, { lineWrap: true });
  const outputLines = output.split('\r\n');

  assert.ok(outputLines.length > 3);
  for (const line of outputLines) assert.ok(line.length <= 120, `超长行: ${line}`);
  assert.ok(output.includes(' _\r\n'));
  assert.equal(output.split('"value11"').length, 2);
  assert.equal(formatVba(output, { lineWrap: true }), output);
});

test('lineWrap 不改写字符串内的逗号，也不拆分带注释的行', () => {
  const longString = `x = Foo("${'a,'.repeat(80)}", 1, 2)`;
  const source = lines('Sub S()', longString, 'End Sub');
  const output = formatVba(source, { lineWrap: true });
  assert.equal(output.includes(`"${'a,'.repeat(80)}"`), true);

  const withComment = `x = Foo(${Array.from({ length: 12 }, (_v, i) => `"v${i}"`).join(', ')}) ' 注释`;
  const commented = lines('Sub S()', withComment, 'End Sub');
  assert.equal(formatVba(commented, { lineWrap: true }), formatVba(commented));
});

test('复杂样例多次格式化结果稳定（幂等）', () => {
  const source = lines(
    '\uFEFFAttribute VB_Name = "Module1"',
    'option explicit',
    '',
    'public sub main()',
    'if a then b = 1: c = 2',
    'select case x',
    'case 1',
    'for i = 1 to 3',
    'with obj',
    's = "  keep  "  \' 注释',
    'end with',
    'next i',
    'case else',
    'x = 1 + _',
    '2',
    'end select',
    'end if',
    'end sub',
  );
  for (const options of [
    {},
    { lineWrap: true },
    { reorderDeclarations: true },
    { formatDesign: true },
    { tabSize: 2 },
    { insertSpaces: false },
    { keywordCase: false },
  ]) {
    const once = formatVba(source, options);
    assert.equal(formatVba(once, options), once, `非幂等: ${JSON.stringify(options)}`);
  }
});
