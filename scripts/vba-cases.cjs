/* eslint-disable */
// VBA 解析器测试用例自动判定（依据 docs/VBA解析器测试用例.md v1.0）。
// 运行前提：npm run compile（使用 dist 编译产物）。输出：output/vba-cases-report.md 与 .json。
'use strict';
const { parseModule } = require('../dist/src/language/parser');
const { VbaService } = require('../dist/src/language/service');
const { writeFileSync } = require('node:fs');

const FAILURE_SEVERITY = 2; // severity <= 2（Error/Warning）计为“报错”；Info/Hint 不计（文档：无错误=零错误零警告）。

// LEX-02 / LEX-02b 生成器：n 个连接符。
const continuation = (n) => {
  const lines = ['Sub T()', '    Dim s As String', '    s = "1"'];
  for (let i = 2; i <= n + 1; i += 1) { lines.push(`        "${i}" & _`); }
  lines.push(`        "${n + 2}"`, 'End Sub');
  return lines.join('\n');
};
const longLine = 'x = ' + '1+'.repeat(510) + '1'; // 1024+ 字符合法语法行

const CASES = [
  // —— LEX ——
  { id: 'LEX-01', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim s As String\n    s = "a" & _\n        "b" & _\n        "c"\nEnd Sub' },
  { id: 'LEX-02', g: 'LEX', kind: 'n', expectLine: 28, note: '25 连接符上限；未实现则漏报', input: continuation(25) },
  { id: 'LEX-02b', g: 'LEX', kind: 'p', input: continuation(24) },
  { id: 'LEX-03', g: 'LEX', kind: 'n', expectLine: 3, note: '续行符后注释；未实现则漏报', input: 'Sub T()\n    Dim s As String\n    s = "a" & _ \' trailing\n        "b"\nEnd Sub' },
  { id: 'LEX-04', g: 'LEX', kind: 'p', note: '应识别两条语句（AST 级未检查）', input: 'Sub T()\n    Dim a As Long, b As Long\n    a = 1: b = 2\nEnd Sub' },
  { id: 'LEX-05', g: 'LEX', kind: 'p', note: 'Remthis 行为以 VBE 实测为准（置信度中）', input: "Rem this is a comment\n' this is also a comment\nSub T()\n    Dim a As Long\n    a = 1 ' trailing\n    a = 2: Rem trailing Rem\nEnd Sub" },
  { id: 'LEX-06', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim Foo As Long\n    foo = 1\n    FOO = 2\nEnd Sub' },
  { id: 'LEX-07', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim [Name] As String\n    [Name] = "x"\n    Dim [Dim] As Long\n    [Dim] = 1\nEnd Sub' },
  { id: 'LEX-08', g: 'LEX', kind: 'p', input: 'Function Foo$()\n    Foo = "x"\nEnd Function\n\nSub T()\n    Dim s$\n    Dim i%\n    Dim l&\n    Dim sng!\n    Dim dbl#\n    Dim cur@\n    s = Foo()\nEnd Sub' },
  { id: 'LEX-09', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim t As String\n    t = Left$("abcdef", 3)\n    t = Mid$("abcdef", 2, 2)\n    t = UCase$("a")\n    t = Trim$("  x  ")\n    t = Chr$(65)\nEnd Sub' },
  { id: 'LEX-10', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim a As Long, b As Double, c As Double, d As Long\n    a = &HFF&\n    b = .5\n    c = 1.5E+3\n    d = 5#\n    b = 5.\nEnd Sub' },
  { id: 'LEX-11', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim d1 As Date, d2 As Date, d3 As Date\n    d1 = #1/1/2020#\n    d2 = #10:00:00 AM#\n    d3 = #May 5, 1993#\nEnd Sub' },
  { id: 'LEX-12', g: 'LEX', kind: 'p', input: 'Sub T()\n    Dim s As String\n    s = "say ""hi"""\nEnd Sub' },
  { id: 'LEX-13', g: 'LEX', kind: 'p', note: 'Line/Name/Time/Get/Seek 非保留字', input: 'Sub T()\n    Dim Name As String\n    Dim Time_1 As Date\n    Name = "x"\n    Dim Line As Long\n    Line = 1\n    Dim Seek_2 As Long\nEnd Sub' },
  { id: 'LEX-14', g: 'LEX', kind: 'p', note: 'Mid 语句/函数双身份', input: 'Sub T()\n    Dim s As String\n    s = "hello"\n    Mid(s, 2, 3) = "XYZ"\n    Dim t As String\n    t = Mid(s, 2, 3)\nEnd Sub' },
  { id: 'LEX-15', g: 'LEX', kind: 'p', note: '解析层不得因 ! 报错', input: 'Sub T()\n    Dim v As Variant\n    v = rs!FieldName.Value\nEnd Sub' },
  { id: 'LEX-16', g: 'LEX', kind: 'n', note: '物理行 1023 字符上限；未实现则漏报', input: 'Sub T()\n' + longLine + '\nEnd Sub' },
  // —— PRE ——
  { id: 'PRE-01', g: 'PRE', kind: 'p', input: 'Option Explicit\nOption Base 1\nOption Compare Text\nOption Private Module\n\nSub T()\nEnd Sub' },
  { id: 'PRE-01b', g: 'PRE', kind: 'n', note: 'Option 位于过程后；未实现则漏报', input: 'Sub T()\nEnd Sub\nOption Explicit' },
  { id: 'PRE-01c', g: 'PRE', kind: 'n', note: 'Option 重复；未实现则漏报', input: 'Option Explicit\nOption Explicit\nSub T()\nEnd Sub' },
  { id: 'PRE-02', g: 'PRE', kind: 'p', note: '未激活分支免检（重点）', input: '#Const DebugMode = True\n\nSub T()\n#If DebugMode Then\n    Dim x As Long\n    x = 1\n#Else\n    This is not valid VBA syntax @@@\n#End If\nEnd Sub' },
  { id: 'PRE-03', g: 'PRE', kind: 'p', input: '#Const A = 1\n\nSub T()\n#If A = 1 Then\n#If Win64 Then\n    Dim p As LongPtr\n#Else\n    Dim p As Long\n#End If\n#End If\nEnd Sub' },
  { id: 'PRE-04', g: 'PRE', kind: 'p', input: 'Sub T()\n#If VBA7 Then\n    Dim a As Long\n#ElseIf Win64 Then\n    Dim b As Long\n#ElseIf Win32 Then\n    Dim c As Long\n#ElseIf Mac Then\n    Dim d As Long\n#End If\nEnd Sub' },
  { id: 'PRE-05', g: 'PRE', kind: 'p', input: 'Attribute VB_Name = "Module1"\nAttribute VB_Description = "test module"\nOption Explicit\n\nSub T()\nEnd Sub' },
  { id: 'PRE-06', g: 'PRE', kind: 'p', input: 'Attribute VB_Name = "Module1"\nOption Explicit\n\nPublic Function Foo() As Long\nAttribute Foo.VB_Description = "returns one"\nAttribute Foo.VB_UserMemId = 0\n    Foo = 1\nEnd Function' },
  { id: 'PRE-07', g: 'PRE', kind: 'p', input: 'Attribute VB_Name = "Module1"\nOption Explicit\n\nPublic gValue As Long\nAttribute gValue.VB_VarDescription = "global value"\nAttribute gValue.VB_VarUserMemId = 0' },
  { id: 'PRE-08', g: 'PRE', kind: 'p', ext: '.cls', input: 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\nAttribute VB_Name = "Class1"\nAttribute VB_PredeclaredId = True\nOption Explicit\n\nPublic Function GiveMeATwo() As Integer\n    GiveMeATwo = 2\nEnd Function' },
  { id: 'PRE-09', g: 'PRE', kind: 'n', note: 'Deftype 位置；未实现则漏报', input: 'Attribute VB_Name = "Module1"\nOption Explicit\n\nDim x As Long\n\nDefInt i-k' },
  { id: 'PRE-10', g: 'PRE', kind: 'p', input: 'Attribute VB_Name = "Module1"\nDefBool b\nDefByte y\nDefInt i-k, N\nDefLng x-z, D\nDefLngLng q\nDefLngPtr w\nDefCur c, l-m\nDefSng a\nDefDbl f\nDefDate t\nDefStr s\nDefObj o\nDefVar v\n\nSub T()\n    Dim icount As Long\n    icount = 1\nEnd Sub' },
  // —— STMT ——
  { id: 'STMT-01', g: 'STMT', kind: 'p', input: 'Public Sub S1()\nEnd Sub\n\nPrivate Sub S2()\nEnd Sub\n\nPublic Static Function F1() As Long\n    F1 = 1\nEnd Function\n\nPrivate Sub S3(ByVal a As Long)\nEnd Sub' },
  { id: 'STMT-01b', g: 'STMT', kind: 'n', options: { moduleKind: 'standard' }, note: '标准模块 Friend；未实现则漏报', input: 'Friend Sub S()\nEnd Sub' },
  { id: 'STMT-02', g: 'STMT', kind: 'p', input: 'Sub T(Optional a As Long = 1, ParamArray b())\n    Dim v As Variant\n    For Each v In b\n    Next v\nEnd Sub' },
  { id: 'STMT-02b', g: 'STMT', kind: 'n', note: 'ParamArray 不在最后；未实现则漏报', input: 'Sub T(ParamArray a(), b As Long)\nEnd Sub' },
  { id: 'STMT-03', g: 'STMT', kind: 'p', note: '调用括号规则', input: 'Sub Foo(a As Long, b As Long)\nEnd Sub\n\nSub T()\n    Call Foo(1, 2)\n    Foo 1, 2\n    Foo (1), 2\nEnd Sub' },
  { id: 'STMT-03b', g: 'STMT', kind: 'n', note: 'Foo (1, 2) 应报错；未实现则漏报', input: 'Sub Foo(a As Long, b As Long)\nEnd Sub\n\nSub T()\n    Foo (1, 2)\nEnd Sub' },
  { id: 'STMT-04', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim a As Long\n    Dim o As Object\n    Let a = 1\n    a = 2\n    Set o = Nothing\nEnd Sub' },
  { id: 'STMT-04b', g: 'STMT', kind: 'n', note: 'a, b = 1 应报错；未实现则漏报', input: 'Sub T()\n    a, b = 1\nEnd Sub' },
  { id: 'STMT-05', g: 'STMT', kind: 'p', note: '单行 If 与冒号（重点）', input: 'Sub T()\n    Dim a As Long, b As Long, x As Long\n    x = 1\n    If x = 1 Then a = 1: b = 2 Else a = 3: b = 4\nEnd Sub' },
  { id: 'STMT-05b', g: 'STMT', kind: 'n', expectLine: 5, input: 'Sub T()\n    Dim x As Long\n    x = 1\n    If x = 1 Then x = 2\nEnd If\nEnd Sub' },
  { id: 'STMT-06', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim x As Long\n    x = 2\n    If x = 1 Then\n        x = 10\n    ElseIf x = 2 Then\n        x = 20\n    ElseIf x = 3 Then\n        x = 30\n    Else\n        x = 40\n    End If\nEnd Sub' },
  { id: 'STMT-07', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim x As Long\n    x = 5\n    Select Case x\n        Case 1 To 5, 8\n            x = 0\n        Case Is > 10\n            x = 1\n        Case 9, Is < 0\n            x = 2\n        Case Else\n            x = 3\n    End Select\nEnd Sub' },
  { id: 'STMT-08', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim i As Long\n    i = 0\n    While i < 10\n        i = i + 1\n    Wend\nEnd Sub' },
  { id: 'STMT-08b', g: 'STMT', kind: 'n', note: 'Exit While 不存在；未实现则漏报', input: 'Sub T()\n    While True\n        Exit While\n    Wend\nEnd Sub' },
  { id: 'STMT-09', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim i As Long\n    i = 0\n    Do While i < 10\n        i = i + 1\n    Loop\n    Do\n        i = i - 1\n    Loop While i > 0\n    Do Until i < -5\n        i = i - 1\n    Loop\n    Do\n        i = i + 1\n        If i > 100 Then Exit Do\n    Loop Until i > 50\nEnd Sub' },
  { id: 'STMT-10', g: 'STMT', kind: 'p', note: '一条 Next 关闭多层（重点）', input: 'Sub T()\n    Dim i As Long, j As Long, total As Long\n    For i = 1 To 3\n        For j = 1 To 3\n            total = total + 1\n        Next j, i\nEnd Sub' },
  { id: 'STMT-11', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim c As Collection\n    Dim v As Variant\n    Set c = New Collection\n    c.Add 1\n    For Each v In c\n    Next v\nEnd Sub' },
  { id: 'STMT-12', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim c As Collection\n    Set c = New Collection\n    With c\n        .Add 1\n        With c\n            .Add 2\n        End With\n        .Add 3\n    End With\nEnd Sub' },
  { id: 'STMT-13', g: 'STMT', kind: 'p', note: 'GoTo/GoSub/行号/标签（重点）', input: 'Sub T()\n    Dim x As Long\n    GoTo 10\n5   x = 1\n    GoTo 30\n10  x = 2\n    GoSub Label1\n    GoTo 30\nLabel1:\n    x = 3\n    Return\n30  x = 4\nEnd Sub' },
  { id: 'STMT-14', g: 'STMT', kind: 'p', note: 'On...GoTo / On...GoSub（重点）', input: 'Sub T()\n    Dim n As Long\n    n = 2\n    On n GoSub Sub1, Sub2\n    On n GoTo Line1, Line2\n    Exit Sub\nSub1:\n    Return\nSub2:\n    Return\nLine1:\n    n = 1\nLine2:\n    n = 2\nEnd Sub' },
  { id: 'STMT-15', g: 'STMT', kind: 'p', note: 'On Error 家族（重点）', input: 'Sub T()\n    On Error GoTo Handler\n    On Error Resume Next\n    On Error GoTo 0\n    On Error GoTo -1\n    Exit Sub\nHandler:\n    Resume Next\n    Resume\n    Resume 10\n10  End\nEnd Sub' },
  { id: 'STMT-16', g: 'STMT', kind: 'p', note: '声明与数组维度', input: 'Sub T()\n    Static c As Long\n    Const k As Long = 5\n    Dim a(10) As Long\n    Dim b(1 To 10, 2 To 20) As Long\n    Dim d(-5 To 5) As Long\n    Dim e() As Long\n    ReDim e(5)\n    ReDim Preserve e(10)\n    Dim f(3) As Variant\n    Erase f\n    Dim g(1 To 3) As Long\n    g(1) = k\nEnd Sub' },
  { id: 'STMT-17', g: 'STMT', kind: 'p', note: '文件 I/O 语句族（重点）', input: 'Sub T()\n    Dim n As Long\n    Dim s As String\n    Dim v As Variant\n    n = FreeFile\n    Open "test.txt" For Output As #n\n    Print #n, "a"; "b", "c"\n    Write #n, "x", 1, 2.5\n    Close #n\n    Open "test.txt" For Input As #n\n    Input #n, s\n    Line Input #n, s\n    Close n\n    Open "test.bin" For Binary As #n\n    Get #n, , v\n    Put #n, 1, v\n    Seek #n, 1\n    Lock #n\n    Unlock #n\n    Close #n\n    Open "test.txt" For Append Access Read Write Shared As #n Len = 128\n    Close #n\n    Name "a.txt" As "b.txt"\n    FileCopy "b.txt", "c.txt"\n    Kill "c.txt"\n    MkDir "d1"\n    ChDir "d1"\n    ChDrive "C"\n    RmDir "d1"\nEnd Sub' },
  { id: 'STMT-18', g: 'STMT', kind: 'p', input: 'Sub T()\n    Beep\n    Randomize\n    Randomize 42\n    DoEvents\n    SendKeys "abc", True\n    AppActivate "Excel"\nEnd Sub' },
  { id: 'STMT-19', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim x As Long\n    x = 1\n    If x = 1 Then Stop\n    End\nEnd Sub' },
  { id: 'STMT-20', g: 'STMT', kind: 'p', input: 'Private Type Point\n    x As Long\n    y As Long\n    Name As String * 8\n    Data(1 To 4) As Double\nEnd Type\n\nSub T()\n    Dim p As Point\n    p.x = 1\nEnd Sub' },
  { id: 'STMT-20b', g: 'STMT', kind: 'n', expectLine: 3, note: '过程体内 Type；未实现则漏报', input: 'Sub T()\n    Dim p As Variant\n    Type Inner\n        x As Long\n    End Type\nEnd Sub' },
  { id: 'STMT-21', g: 'STMT', kind: 'p', input: 'Public Enum Color\n    Red\n    Green = 5\n    Blue = &HFF\n    Dark = -1\n    Next_ = Blue + 1\nEnd Enum' },
  { id: 'STMT-22', g: 'STMT', kind: 'p', input: 'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)\nPublic Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As LongPtr\nPrivate Declare Sub OldWay Lib "legacy" (x As Any)' },
  { id: 'STMT-23', g: 'STMT', kind: 'p', ext: '.cls', input: 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = "Class1"\nOption Explicit\n\nImplements Collection\n\nPublic Event Changed(ByVal NewValue As Variant)\nPrivate WithEvents mSource As Class2\n\nPrivate Sub Class_Initialize()\nEnd Sub\n\nPrivate Sub Class_Terminate()\nEnd Sub\n\nPublic Default Property Get Item(ByVal i As Long) As Variant\nAttribute Item.VB_UserMemId = 0\n    Item = i\nEnd Property\n\nFriend Sub InternalHelper()\nEnd Sub' },
  { id: 'STMT-24', g: 'STMT', kind: 'p', input: 'Sub T()\n    Dim x As Long\n    GoTo L\nL: x = 1\nEnd Sub' },
  // —— EXPR ——
  { id: 'EXPR-01', g: 'EXPR', kind: 'p', note: '优先级全表（AST 检查未做，仅不误报）', input: 'Sub T()\n    Dim a As Long\n    a = 1 Or 2 Xor 3 Eqv 4 Imp 5\n    a = 6 \\ 2 Mod 2\n    Dim s As String\n    s = "a" & "b" & "c"\nEnd Sub' },
  { id: 'EXPR-02', g: 'EXPR', kind: 'p', note: '-2^2 结合性（AST 检查未做）', input: 'Sub T()\n    Debug.Print -2 ^ 2\nEnd Sub' },
  { id: 'EXPR-03', g: 'EXPR', kind: 'p', note: '命名参数与空实参位', input: 'Sub Foo(ByVal a As Long, ByVal b As Long, ByVal c As Long)\nEnd Sub\n\nSub T()\n    Foo a:=1, b:=2, c:=3\n    Foo 1, c:=3, b:=2\n    Foo , 2, 3\n    Foo 1, , 3\nEnd Sub' },
  { id: 'EXPR-03b', g: 'EXPR', kind: 'n', note: '位置实参跟在命名后；未实现则漏报', input: 'Sub Foo(ByVal a As Long, ByVal b As Long, ByVal c As Long)\nEnd Sub\n\nSub T()\n    Foo 1, b:=2, 3\nEnd Sub' },
  { id: 'EXPR-04', g: 'EXPR', kind: 'p', input: 'Sub T()\n    Dim b As Boolean\n    b = "abc" Like "a?c"\n    b = "abc" Like "a*"\n    b = "a5" Like "a#"\n    b = "a" Like "[a-c]"\nEnd Sub' },
  { id: 'EXPR-05', g: 'EXPR', kind: 'p', input: 'Sub T()\n    Dim o As Object\n    Set o = New Collection\n    If TypeOf o Is Collection Then\n        Debug.Print "yes"\n    End If\nEnd Sub' },
  { id: 'EXPR-06', g: 'EXPR', kind: 'p', input: 'Sub T()\n    Dim c As Collection\n    Set c = New Collection\n    Dim d As New Collection\n    d.Add 1\nEnd Sub' },
  { id: 'EXPR-07', g: 'EXPR', kind: 'p', input: 'Sub T()\n    Dim v As Variant\n    v = a!b.c\n    v = a!b!c\n    v = a.b!c.d\nEnd Sub' },
  { id: 'EXPR-08', g: 'EXPR', kind: 'p', input: 'Public Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As Long) As Long\n\nSub T()\n    Dim r As Long\n    r = EnumWindows(AddressOf Callback, 0)\nEnd Sub\n\nPublic Function Callback(ByVal hWnd As Long, ByVal lParam As Long) As Long\n    Callback = 1\nEnd Function' },
  { id: 'EXPR-08b', g: 'EXPR', kind: 'n', note: 'AddressOf 赋值右侧；未实现则漏报', input: 'Public Function Callback() As Long\n    Callback = 1\nEnd Function\n\nSub T()\n    Dim p As LongPtr\n    p = AddressOf Callback\nEnd Sub' },
  { id: 'EXPR-09', g: 'EXPR', kind: 'p', note: '左值 a.b(c).d(e) = f', input: 'Sub T()\n    a.b(c).d(e) = f\nEnd Sub' },
  { id: 'EXPR-10', g: 'EXPR', kind: 'p', input: 'Sub T()\n    Dim a(5) As Long\n    a(1) = 2\n    Dim x As Long\n    x = a(1)\n    Debug.Print (a(1))\nEnd Sub' },
  { id: 'EXPR-11', g: 'EXPR', kind: 'n', note: '元组不支持；未实现则漏报', input: 'Sub T()\n    a = (1, 2)\nEnd Sub' },
  // —— SEM ——
  { id: 'SEM-01a', g: 'SEM', kind: 'n', note: 'Option Explicit 未声明', input: 'Option Explicit\n\nSub T()\n    x = 1\nEnd Sub' },
  { id: 'SEM-01b', g: 'SEM', kind: 'p', input: 'Sub T()\n    x = 1\nEnd Sub' },
  { id: 'SEM-02', g: 'SEM', kind: 'p', input: 'Dim g As Long\n\nSub T()\n    Dim g As Long\n    g = 1\nEnd Sub' },
  { id: 'SEM-03', g: 'SEM', kind: 'n', note: '重复定义；未实现则漏报', input: 'Sub Foo()\nEnd Sub\n\nSub Foo()\nEnd Sub' },
  { id: 'SEM-04', g: 'SEM', kind: 'p', input: 'Sub T()\n    Helper\nEnd Sub\n\nSub Helper()\nEnd Sub' },
  { id: 'SEM-05', g: 'SEM', kind: 'n', note: '无重载；未实现则漏报', input: 'Sub Foo(ByVal a As Long)\nEnd Sub\n\nSub Foo(ByVal a As String)\nEnd Sub' },
  { id: 'SEM-06-1', g: 'SEM', kind: 'n', note: 'Optional 在必选前；未实现则漏报', input: 'Sub T(Optional a As Long, b As Long)\nEnd Sub' },
  { id: 'SEM-06-2', g: 'SEM', kind: 'n', note: '数组参数须 ByRef；未实现则漏报', input: 'Sub T(ByVal a() As Long)\nEnd Sub' },
  { id: 'SEM-06-3', g: 'SEM', kind: 'n', note: 'ParamArray 不在最后；未实现则漏报', input: 'Sub T(ParamArray a() As Variant, b As Long)\nEnd Sub' },
  { id: 'SEM-06-4', g: 'SEM', kind: 'n', note: '两个 ParamArray；未实现则漏报', input: 'Sub T(ParamArray a(), ParamArray b())\nEnd Sub' },
  { id: 'SEM-07', g: 'SEM', kind: 'p', input: 'Private mVal As Variant\n\nPublic Property Get Item(ByVal i As Long) As Variant\n    Item = mVal\nEnd Property\n\nPublic Property Let Item(ByVal i As Long, ByVal v As Variant)\n    mVal = v\nEnd Property' },
  { id: 'SEM-07b', g: 'SEM', kind: 'n', ext: '.cls', note: '两个 Default 属性；未实现则漏报', input: 'Public Default Property Get A() As Variant\n    A = 1\nEnd Property\n\nPublic Default Property Get B() As Variant\n    B = 2\nEnd Property' },
  { id: 'SEM-08', g: 'SEM', kind: 'n', ext: '.cls', options: { moduleKind: 'class' }, note: '类模块 Public Type；未实现则漏报', input: 'Public Type T\n    x As Long\nEnd Type' },
  { id: 'SEM-09', g: 'SEM', kind: 'p', input: 'Const k As Long = 1 + 2 * 3\nConst s As String = "a" & "b"\n\nSub T()\nEnd Sub' },
  { id: 'SEM-09b', g: 'SEM', kind: 'n', note: '常量须为常量表达式；未实现则漏报', input: 'Function SomeFunction() As Long\n    SomeFunction = 1\nEnd Function\n\nConst k As Long = SomeFunction()' },
  { id: 'SEM-10', g: 'SEM', kind: 'p', input: 'Function Fact(ByVal n As Long) As Long\n    If n <= 1 Then\n        Fact = 1\n    Else\n        Fact = n * Fact(n - 1)\n    End If\nEnd Function' },
  { id: 'SEM-11', g: 'SEM', kind: 'n', ext: '.cls', options: { moduleKind: 'class' }, note: 'WithEvents 须对象类型；未实现则漏报', input: 'Private WithEvents x As Long' },
  // —— VBA ——
  { id: 'VBA-01', g: 'VBA', kind: 'p', input: 'Sub T()\n10  Dim x As Long\n    x = 1\n20  x = 2\n    Debug.Print Erl\nEnd Sub' },
  { id: 'VBA-04', g: 'VBA', kind: 'p', input: 'Global gValue As Long\nGlobal Const gk As Long = 1' },
  { id: 'VBA-05', g: 'VBA', kind: 'p', note: 'Attribute 变体容忍', input: 'Attribute VB_Name = "M"\nPublic Function NewEnum() As Variant\nAttribute NewEnum.VB_MemberFlags = "40"\n    NewEnum = 0\nEnd Function\n\nSub T()\nAttribute T.VB_ProcData.VB_Invoke_Func = "\\n14"\nEnd Sub' },
  { id: 'VBA-06', g: 'VBA', kind: 'p', input: '#If VBA7 Then\n    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)\n#Else\n    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\n#End If' },
  { id: 'VBA-08', g: 'VBA', kind: 'p', input: 'Private Type T1\n    a As Long\nEnd Type\n\nPrivate Type T2\n    b As Long\nEnd Type\n\nSub T()\n    Dim s As String * 10\n    LSet s = "abc"\n    RSet s = "abc"\n    Dim x As T1, y As T2\n    LSet x = y\nEnd Sub' },
  { id: 'VBA-11', g: 'VBA', kind: 'p', input: 'Option Compare Database\n\nSub T()\nEnd Sub' },
  // —— DIAG（自定义判定）——
  { id: 'DIAG-01', g: 'DIAG', kind: 'custom', maxErrors: 3, note: '雪崩抑制：错误总数 ≤ 3', input: 'Sub T()\n    If x = 1 Then\n        For i = 1 To 10\n            With obj\n                .Method\nEnd Sub' },
  { id: 'DIAG-02', g: 'DIAG', kind: 'custom', needLines: [3, 4], note: '两处错误均报出', input: 'Sub T()\n    Dim a As Long\n    a = = 1\n    Dim\n    b = 2\nEnd Sub' },
  { id: 'DIAG-03', g: 'DIAG', kind: 'method', note: '消息对齐：人工/后续项' },
  { id: 'DIAG-04', g: 'DIAG', kind: 'p', note: '运行期错误不得静态报出', input: 'Sub T()\n    Dim i As Integer\n    Dim total As Long\n    For i = 1 To 10\n        total = total + i\n    Next i\n    Select Case total\n        Case 1\n            total = 2\n    End Select\nEnd Sub' },
  { id: 'DIAG-05', g: 'DIAG', kind: 'n', note: '@@@ 非法；未实现则漏报', input: 'Sub T()\n    Dim s As String\n    s = "a" & _\n        "b" & @@@\nEnd Sub' },
  { id: 'DIAG-06', g: 'DIAG', kind: 'custom', minErrors: 4, note: '5 处独立错误至少报 4', input: 'Sub T()\n    a = = 1\n    Dim\n    If x > 1\n    For i = 1\n    Select Case y\nEnd Sub' },
  // —— ENG ——
  { id: 'ENG-01-1', g: 'ENG', kind: 'p', ext: '.bas', input: 'Attribute VB_Name = "Module1"\nSub T()\nEnd Sub' },
  { id: 'ENG-01-2', g: 'ENG', kind: 'p', ext: '.cls', input: 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = "Class1"\nOption Explicit\n\nSub T()\nEnd Sub' },
  { id: 'ENG-01-3', g: 'ENG', kind: 'p', ext: '.frm', note: '设计段跳过后解析代码段', input: 'VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} \n   Caption         =   "UserForm1"\n   ClientHeight    =   3180\n   OleObjectBlob   =   "UserForm1.frx":0000\nEnd\nAttribute VB_Name = "UserForm1"\nAttribute VB_GlobalNameSpace = False\nPrivate Sub CommandButton1_Click()\n    MsgBox "ok"\nEnd Sub' },
  { id: 'ENG-01-4', g: 'ENG', kind: 'method', note: 'ANSI/UTF-16：按需求 INP-02 设计为提示并暂停语义分析，不解析' },
  { id: 'ENG-02', g: 'ENG', kind: 'custom', cross: true, note: '多模块交叉引用', inputA: 'Public Sub Shared1()\nEnd Sub\n\nPublic Enum E\n    One\nEnd Enum', inputB: 'Sub T()\n    Shared1\n    Dim x As E\nEnd Sub' },
  { id: 'ENG-03', g: 'ENG', kind: 'method', note: 'AST 保留 Let/Set 与注释：人工项' },
  { id: 'ENG-04', g: 'ENG', kind: 'method', note: '符号表导出：由 AI 工具/引用查询承担' },
  { id: 'ENG-05', g: 'ENG', kind: 'method', note: '性能基准：见 output/lsp-benchmark.json 与 benchmark' },
];

function runCase(c) {
  const uri = `file:///case${c.ext ?? '.bas'}`;
  if (c.cross) {
    const modA = parseModule('file:///engA.bas', c.inputA);
    const modB = parseModule('file:///engB.bas', c.inputB);
    const service = new VbaService({ modules: [modA, modB], typeIndexes: [] });
    const call = service.tokenAt(modB, { line: 1, character: 8 });
    const symbol = call && service.resolve(modB, call);
    const errorsB = modB.diagnostics.filter(d => d.severity <= FAILURE_SEVERITY);
    const resolved = !!symbol && symbol.uri === modA.uri;
    return { id: c.id, group: c.g, kind: c.kind, verdict: resolved && errorsB.length === 0 ? '通过' : '失败', note: c.note, detail: resolved ? '' : 'Shared1 未解析到模块 A', errorLines: [], messages: [] };
  }
  const module = parseModule(uri, c.input, c.options ?? {});
  // 语义诊断（undeclared / name-case / type-mismatch / argument-count）来自 VbaService，须与结构诊断合并判定。
  const service = new VbaService({ modules: [module], typeIndexes: [] });
  const all = [...module.diagnostics, ...service.diagnostics(module)];
  const failing = all.filter(d => d.severity <= FAILURE_SEVERITY);
  const errorLines = [...new Set(failing.map(d => d.range.start.line + 1))].sort((a, b) => a - b);
  let verdict;
  let detail = '';
  if (c.kind === 'p') {
    verdict = failing.length === 0 ? '通过' : '失败（误报）';
    if (failing.length) { detail = failing.slice(0, 3).map(d => `L${d.range.start.line + 1} ${d.message}`); }
  } else if (c.kind === 'n') {
    const hit = failing.some(d => c.expectLine === undefined || d.range.start.line + 1 === c.expectLine);
    verdict = hit ? '通过' : '失败（漏报）';
    if (!hit) { detail = errorLines.length ? `实际报错行：${errorLines.join(',')}` : '未报 Error/Warning'; }
  } else if (c.kind === 'custom') {
    if (c.maxErrors !== undefined) { verdict = failing.length <= c.maxErrors ? '通过' : '失败'; detail = `实际 ${failing.length} 条：${errorLines.join(',')}`; }
    else if (c.needLines) { const missing = c.needLines.filter(line => !errorLines.includes(line)); verdict = missing.length === 0 ? '通过' : '失败（漏报）'; detail = missing.length ? `未报行：${missing.join(',')}` : ''; }
    else if (c.minErrors !== undefined) { verdict = errorLines.length >= c.minErrors ? '通过' : '失败（漏报）'; detail = `实际报错行：${errorLines.join(',')}`; }
  }
  return { id: c.id, group: c.g, kind: c.kind, verdict, note: c.note, detail, errorLines, messages: failing.slice(0, 3).map(d => d.message) };
}

const skipped = CASES.filter(c => c.kind === 'method').map(c => ({ id: c.id, group: c.g, note: c.note }));
const results = CASES.filter(c => c.kind !== 'method').map(runCase);

const groups = ['LEX', 'PRE', 'STMT', 'EXPR', 'SEM', 'VBA', 'DIAG', 'ENG'];
const summary = groups.map(g => {
  const rows = results.filter(r => r.group === g);
  return { group: g, total: rows.length, pass: rows.filter(r => r.verdict === '通过').length, falsePositive: rows.filter(r => r.verdict.includes('误报')).length, miss: rows.filter(r => r.verdict.includes('漏报')).length, otherFail: rows.filter(r => r.verdict === '失败').length };
});
const overall = { total: results.length, pass: results.filter(r => r.verdict === '通过').length, fail: results.filter(r => r.verdict !== '通过').length };
const report = { generatedAt: new Date().toISOString(), parser: 'VBAdminx dist/src/language/parser（容错结构级分析器，非完整 VBA 编译器）', overall, summary, results, skipped };
writeFileSync('output/vba-cases-report.json', JSON.stringify(report, null, 2) + '\n');

const lines = ['# VBA 解析器测试用例检查报告', '', `生成时间：${report.generatedAt}`, '', `被测对象：VBAdminx 解析器（容错结构级分析器，非完整 VBA 编译器）。判定口径：正例零 Error/Warning；反例须报 Error 且行号匹配（文档指定行时）。Information/Hint 不计入失败。`, '', `**总计 ${overall.total} 项自动判定：通过 ${overall.pass}，失败 ${overall.fail}**。另有 ${skipped.length} 项方法类用例未自动判定（见文末）。`, '', '| 组 | 判定项 | 通过 | 误报 | 漏报 | 其他失败 |', '|---|---:|---:|---:|---:|---:|'];
for (const s of summary) { lines.push(`| ${s.group} | ${s.total} | ${s.pass} | ${s.falsePositive} | ${s.miss} | ${s.otherFail} |`); }
lines.push('', '## 逐项结果', '');
for (const g of groups) {
  const rows = results.filter(r => r.group === g);
  if (!rows.length) continue;
  lines.push(`### ${g}`, '', '| 用例 | 判定 | 实际输出摘要 | 备注 |', '|---|---|---|---|');
  for (const r of rows) {
    const summaryText = [r.verdict, r.detail, ...r.messages.map(m => `“${m}”`)].filter(Boolean).join('；').replace(/\|/g, '\\|');
    lines.push(`| ${r.id} | ${r.verdict} | ${summaryText || '—'} | ${r.note ?? ''} |`);
  }
  lines.push('');
}
lines.push('## 方法类用例（未自动判定）', '', '| 用例 | 说明 |', '|---|---|');
for (const s of skipped) { lines.push(`| ${s.id} | ${s.note} |`); }
lines.push('', '## 主要缺口解读', '', '1. 语法级限制：行连接上限/续行后注释、过程体内 Type、调用括号规则、元组、位置实参次序等需要真正的语法校验；当前解析器为容错结构级分析，不做完整语法校验，相关反例漏报属预期范围。', '2. 语义级限制：重复定义、重载、Property 配对、WithEvents 类型、常量表达式、Option 位置、Deftype 位置等需要符号级语义校验，尚未实现。', '3. 表达式层：除条件编译表达式外不做表达式解析与 AST 建模，EXPR 组正例“不误报”属被动通过，反例漏报。', '4. 通过的用例也不代表完整支持：部分正例仅因解析器未深入校验而未报错（见各用例备注）。');
writeFileSync('output/vba-cases-report.md', lines.join('\n') + '\n');
console.log(JSON.stringify({ overall, summary }, null, 2));
