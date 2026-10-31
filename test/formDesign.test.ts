import assert from 'node:assert/strict';
import test from 'node:test';
import {
  controlBold,
  controlCaption,
  controlFontName,
  controlFontSize,
  oleColorToCss,
  parseFormDesign,
  parseNumericProperty,
  parsePropertyValue,
  resolveControlType,
} from '../src/formDesign';
import { buildRenderPayload, formClientSize, renderFormBody, renderFormPage } from '../src/formHtml';
import { inferFormDesignFromCode } from '../src/formInfer';

const FORM = `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Demo &Form"
   ClientHeight    =   3000
   ClientWidth     =   4500
   BackColor       =   &H8000000F&
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} CommandButton1
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      Top             =   180
      Width           =   1215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Weight          =   700
      EndProperty
   End
   Begin VB.Label Label1
      Caption         =   "A && B &C"
      Height          =   300
      Left            =   120
      Top             =   780
      Width           =   2000
   End
   Begin {11111111-2222-3333-4444-555555555555} Mystery1
      Left            =   0
      Top             =   0
      Width           =   100
      Height          =   100
   End
End
Attribute VB_Name = "UserForm1"
`;

test('解析窗体根节点与客户区尺寸', () => {
  const design = parseFormDesign(FORM);
  assert.ok(design.root);
  assert.equal(design.root.type, 'UserForm');
  assert.equal(design.root.name, 'UserForm1');
  assert.equal(design.root.properties['ClientWidth'], '4500');
  assert.equal(design.root.children.length, 3);
  assert.deepEqual(formClientSize(design), { width: 4500, height: 3000 });
});

test('GUID 与 VB.Xxx 控件类型归一化', () => {
  assert.equal(resolveControlType('{D7053240-CE69-11CD-A777-00DD01143C57}'), 'CommandButton');
  assert.equal(resolveControlType('{d7053240-ce69-11cd-a777-00dd01143c57}'), 'CommandButton');
  assert.equal(resolveControlType('{8BD21D10-EC42-11CE-9E0D-00AA006002F3}'), 'TextBox');
  assert.equal(resolveControlType('{6E182020-F460-11CE-9BCD-00AA00608E01}'), 'Frame');
  assert.equal(resolveControlType('{EAE50EB0-4A62-11CE-BED6-00AA00611080}'), 'TabStrip');
  assert.equal(resolveControlType('VB.Label'), 'Label');
  assert.equal(resolveControlType('MSForms.TextBox'), 'TextBox');
  assert.match(resolveControlType('{11111111-2222-3333-4444-555555555555}'), /^Unknown/);
});

test('BeginProperty Font 被折叠为 Font.* 属性', () => {
  const button = parseFormDesign(FORM).root!.children[0]!;
  assert.equal(button.type, 'CommandButton');
  assert.equal(button.properties['Font.Name'], 'Tahoma');
  assert.equal(button.properties['Font.Size'], '9');
  assert.equal(button.properties['Font.Weight'], '700');
  assert.equal(controlFontName(button), 'Tahoma');
  assert.equal(controlFontSize(button), 9);
  assert.equal(controlBold(button), true);
});

test('字符串属性与访问键处理', () => {
  assert.equal(parsePropertyValue('"He said ""hi"""'), 'He said "hi"');
  assert.equal(parsePropertyValue('495  \'twips'), '495');
  assert.equal(parsePropertyValue('&H8000000F&'), '&H8000000F&');
  const label = parseFormDesign(FORM).root!.children[1]!;
  assert.equal(controlCaption(label), 'A & B C');
});

test('OLE 颜色与数值解析', () => {
  assert.equal(parseNumericProperty('&H8000000F&'), 0x8000000f);
  assert.equal(parseNumericProperty('-1'), -1);
  assert.equal(parseNumericProperty('abc', 7), 7);
  assert.equal(oleColorToCss('&H8000000F&', '#000'), '#f0f0f0');
  assert.equal(oleColorToCss('&H0000FF&', '#000'), 'rgb(255, 0, 0)');
  assert.equal(oleColorToCss('&HFF0000&', '#000'), 'rgb(0, 0, 255)');
  assert.equal(oleColorToCss(undefined, '#123'), '#123');
});

test('渲染体包含控件与文本，未知控件以占位呈现', () => {
  const design = parseFormDesign(FORM);
  const body = renderFormBody(design, 'empty');
  assert.match(body, /vba-button/);
  assert.match(body, /OK/);
  assert.match(body, /vba-label/);
  assert.match(body, /A &amp; B C/);
  assert.match(body, /vba-unknown/);
  assert.match(body, /Unknown/);
});

test('完整页面带 CSP nonce 与热更新脚本', () => {
  const design = parseFormDesign(FORM);
  const page = renderFormPage(design, { nonce: 'abc123', title: 'Form preview: Demo', hint: 'hint', emptyText: 'no design' });
  assert.match(page, /nonce-abc123/);
  assert.match(page, /Content-Security-Policy/);
  assert.match(page, /type !== 'render'/);
  const payload = buildRenderPayload(design, 'Form preview: Demo', 'no design');
  assert.ok(payload.width > 0 && payload.height > 0);
  assert.equal(payload.caption, 'Demo Form');
});

test('无设计段时返回空模型', () => {
  const design = parseFormDesign('Attribute VB_Name = "Module1"\nSub Main()\nEnd Sub\n');
  assert.equal(design.root, undefined);
  assert.equal(renderFormBody(design, 'no design'), '<div class="empty">no design</div>');
});

const CODE_ONLY_FORM = `' 窗体变量
Public SelectedSheetName As String

Private Sub UserForm_Initialize()
    With Me
        .Width = 300
        .Height = 150
        .Caption = "请选择要处理的工作表"
        With .lblTitle
            .Caption = "请选择要处理的工作表"
            .Top = 10
            .Left = 10
            .Width = 280
            .Height = 16
            .Font.Bold = False
            .Font.Name = "微软雅黑"
            .Font.Size = 14
        End With
        With .cmbSheets
            .Top = 30
            .Left = 10
            .Width = 280
            .Height = 20
        End With
        With .btnOK
            .Caption = "确定"
            .Top = 70
            .Left = 70
            .Width = 70
            .Height = 24
            .Default = True
        End With
        With .btnCancel
            .Caption = "取消"
            .Top = 70
            .Left = 160
            .Width = 70
            .Height = 24
        End With
    End With
End Sub
`;

test('无设计段时按初始化代码推断窗体', () => {
  assert.equal(parseFormDesign(CODE_ONLY_FORM).root, undefined);
  const design = inferFormDesignFromCode(CODE_ONLY_FORM, 'frmSelectSheet');
  assert.ok(design.root);
  assert.equal(design.inferred, true);
  assert.equal(design.root.properties['ClientWidth'], '6000');
  assert.equal(design.root.properties['ClientHeight'], '3000');
  const byName = new Map(design.root.children.map(child => [child.name, child]));
  assert.equal(byName.get('lblTitle')!.type, 'Label');
  assert.equal(byName.get('cmbSheets')!.type, 'ComboBox');
  assert.equal(byName.get('btnOK')!.type, 'CommandButton');
  assert.equal(byName.get('btnOK')!.properties['Caption'], '确定');
  assert.equal(byName.get('btnOK')!.properties['Left'], '1400');
  assert.equal(byName.get('btnOK')!.properties['Width'], '1400');
  assert.equal(byName.get('lblTitle')!.properties['Top'], '200');
  assert.equal(byName.get('lblTitle')!.properties['Width'], '5600');
  assert.equal(byName.get('lblTitle')!.properties['FontBold'], 'False');
  assert.equal(byName.get('lblTitle')!.properties['FontName'], '微软雅黑');
  assert.equal(byName.get('lblTitle')!.properties['FontSize'], '14');
});

test('只有事件过程的窗体也能生成默认布局', () => {
  const code = [
    'Private Sub btnOK_Click()',
    'End Sub',
    'Private Sub btnCancel_Click()',
    'End Sub',
    'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)',
    'End Sub',
  ].join('\r\n');
  const design = inferFormDesignFromCode(code, 'frmSelectSheetV2');
  assert.ok(design.root);
  assert.equal(design.inferred, true);
  assert.equal(design.root.children.length, 2);
  for (const child of design.root.children) {
    assert.equal(child.type, 'CommandButton');
    assert.ok(child.properties['Left'] && child.properties['Top']);
    assert.ok(child.properties['Width'] && child.properties['Height']);
  }
});

test('没有设计段且没有可用布局信息时仍返回空模型', () => {
  assert.equal(inferFormDesignFromCode('Option Explicit\r\nPublic Sub Main()\r\nEnd Sub\r\n').root, undefined);
});
