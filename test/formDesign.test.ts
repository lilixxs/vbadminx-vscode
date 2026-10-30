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
