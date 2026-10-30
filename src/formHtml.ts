/**
 * 把解析后的窗体设计渲染成 HTML（不依赖 vscode 模块）。
 *
 * 坐标单位是 VBA 的 twip（1/1440 英寸），按 96 DPI 换算为像素（1 px = 15 twips）。
 * 渲染是“近似还原”：只表达位置、尺寸、常用样式和文本，不执行任何 VBA 代码。
 */
import {
  controlBold,
  controlCaption,
  controlEnabled,
  controlFontName,
  controlFontSize,
  controlItalic,
  controlVisible,
  oleColorToCss,
  parseNumericProperty,
  parsePropertyValue,
  type FormControl,
  type FormDesign,
} from './formDesign';

const TWIPS_PER_PIXEL = 15;

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function twipsToPx(value: number): number {
  return Math.round(value / TWIPS_PER_PIXEL);
}

function prop(control: FormControl, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = control.properties[name];
    if (value !== undefined) { return value; }
  }
  return undefined;
}

function numeric(control: FormControl, names: string[], fallback = 0): number {
  return parseNumericProperty(prop(control, ...names), fallback);
}

function isTrue(value: string | undefined): boolean {
  return value === 'True' || value === '-1';
}

function fontSizePx(control: FormControl): number {
  return Math.max(6, Math.round(controlFontSize(control) * (96 / 72)));
}

function fontStyle(control: FormControl): string {
  return `font-family:${escapeHtml(controlFontName(control))},Tahoma,sans-serif;font-size:${fontSizePx(control)}px;`
    + `font-weight:${controlBold(control) ? 700 : 400};font-style:${controlItalic(control) ? 'italic' : 'normal'};`;
}

function foreground(control: FormControl, fallback = '#111111'): string {
  return oleColorToCss(prop(control, 'ForeColor'), fallback);
}

function background(control: FormControl, fallback: string): string {
  return oleColorToCss(prop(control, 'BackColor'), fallback);
}

function alignment(control: FormControl): string {
  const value = numeric(control, ['Alignment'], 0);
  return value === 1 ? 'right' : value === 2 ? 'center' : 'left';
}

function textValue(control: FormControl): string {
  return prop(control, 'Text', 'Value') ?? '';
}

function box(control: FormControl, className: string, inner: string, extra = ''): string {
  const left = numeric(control, ['Left']);
  const top = numeric(control, ['Top']);
  const width = numeric(control, ['Width'], 0);
  const height = numeric(control, ['Height'], 0);
  const style: string[] = [
    `left:${twipsToPx(left)}px`,
    `top:${twipsToPx(top)}px`,
    width > 0 ? `width:${twipsToPx(width)}px` : 'width:auto',
    height > 0 ? `height:${twipsToPx(height)}px` : 'height:auto',
    `color:${foreground(control)}`,
    fontStyle(control),
    extra,
  ];
  if (!controlVisible(control)) { style.push('display:none'); }
  return `<div class="control ${className}" style="${style.join(';')}">${inner}</div>`;
}

function captionHtml(control: FormControl): string {
  return escapeHtml(controlCaption(control));
}

function checkState(control: FormControl): boolean {
  return isTrue(prop(control, 'Value'));
}

function renderTextBox(control: FormControl): string {
  const text = escapeHtml(textValue(control));
  const align = alignment(control);
  return box(control, 'vba-textbox', `<span>${text}</span>`, `background:${background(control, '#ffffff')};text-align:${align};`);
}

function renderComboBox(control: FormControl): string {
  const text = escapeHtml(textValue(control));
  return box(control, 'vba-combobox', `<span class="grow">${text}</span><span class="arrow">▾</span>`, `background:${background(control, '#ffffff')};`);
}

function renderListBox(control: FormControl): string {
  return box(control, 'vba-listbox', '<div class="row"></div><div class="row"></div><div class="row"></div>', `background:${background(control, '#ffffff')};`);
}

function renderButton(control: FormControl): string {
  const pressed = control.type === 'ToggleButton' && checkState(control);
  const disabled = controlEnabled(control) ? '' : ' disabled';
  return box(control, `vba-button${pressed ? ' pressed' : ''}${disabled}`, `<span>${captionHtml(control)}</span>`, `background:${background(control, pressed ? '#d6d6d6' : '#f0f0f0')};`);
}

function renderCheck(control: FormControl): string {
  const checked = checkState(control);
  return box(control, 'vba-check', `<span class="mark">${checked ? '✓' : ''}</span><span class="grow">${captionHtml(control)}</span>`, 'background:transparent;');
}

function renderOption(control: FormControl): string {
  const checked = checkState(control);
  return box(control, 'vba-option', `<span class="mark">${checked ? '●' : ''}</span><span class="grow">${captionHtml(control)}</span>`, 'background:transparent;');
}

function renderScrollBar(control: FormControl): string {
  const width = numeric(control, ['Width']);
  const height = numeric(control, ['Height']);
  const vertical = height >= width;
  return box(control, `vba-scroll ${vertical ? 'vertical' : 'horizontal'}`, '<div class="thumb"></div>', `background:${background(control, '#e6e6e6')};`);
}

function renderSpin(control: FormControl): string {
  return box(control, 'vba-spin', '<div class="up">▴</div><div class="down">▾</div>', `background:${background(control, '#f0f0f0')};`);
}

function renderImage(control: FormControl): string {
  return box(control, 'vba-image', '<span>Image</span>', 'background:#dcdcdc;');
}

function renderLine(control: FormControl): string {
  return box(control, 'vba-line', '', 'background:#7a7a7a;');
}

function renderContainer(control: FormControl, children: string, tabs = false): string {
  const caption = captionHtml(control);
  const header = tabs
    ? `<div class="vba-tab">${caption || '&nbsp;'}</div>`
    : caption ? `<div class="vba-frame-caption">${caption}</div>` : '';
  return box(control, `vba-container${tabs ? ' tabs' : ''}`, `${header}<div class="vba-client">${children}</div>`, `background:${background(control, '#f0f0f0')};`);
}

function renderUnknown(control: FormControl): string {
  return box(control, 'vba-unknown', `<span>${escapeHtml(control.type)}</span>`, 'border:1px dashed #999;background:transparent;');
}

function renderControl(control: FormControl): string {
  const children = control.children.map(renderControl).join('');
  switch (control.type) {
    case 'CommandButton':
    case 'ToggleButton':
      return renderButton(control);
    case 'Label':
      return box(control, 'vba-label', `<span>${captionHtml(control)}</span>`, `background:${background(control, 'transparent')};text-align:${alignment(control)};`);
    case 'TextBox':
      return renderTextBox(control);
    case 'ComboBox':
      return renderComboBox(control);
    case 'ListBox':
      return renderListBox(control);
    case 'CheckBox':
      return renderCheck(control);
    case 'OptionButton':
      return renderOption(control);
    case 'ScrollBar':
      return renderScrollBar(control);
    case 'SpinButton':
      return renderSpin(control);
    case 'Image':
      return renderImage(control);
    case 'Line':
      return renderLine(control);
    case 'Frame':
      return renderContainer(control, children);
    case 'TabStrip':
    case 'MultiPage':
      return renderContainer(control, children, true);
    case 'UserForm':
      return children;
    default:
      return renderUnknown(control);
  }
}

/** 渲染窗体内部（stage）的 HTML。 */
export function renderFormBody(model: FormDesign, emptyText = ''): string {
  const root = model.root;
  if (!root || root.children.length === 0) {
    return emptyText ? `<div class="empty">${escapeHtml(emptyText)}</div>` : '';
  }
  return root.children.map(renderControl).join('');
}

/** 计算窗体客户区尺寸（twips），缺失时按子控件范围推断。 */
export function formClientSize(model: FormDesign): { width: number; height: number } {
  const root = model.root;
  if (!root) { return { width: 4560, height: 3015 }; }
  let width = numeric(root, ['ClientWidth', 'Width'], 0);
  let height = numeric(root, ['ClientHeight', 'Height'], 0);
  if (width <= 0 || height <= 0) {
    for (const child of root.children) {
      width = Math.max(width, numeric(child, ['Left']) + numeric(child, ['Width']));
      height = Math.max(height, numeric(child, ['Top']) + numeric(child, ['Height']));
    }
  }
  return { width: width > 0 ? width : 4560, height: height > 0 ? height : 3015 };
}

export interface FormPageOptions {
  nonce: string;
  title: string;
  hint: string;
  emptyText: string;
}

/** 生成完整的 WebView 文档；热更新时用 postMessage 替换 #stage。 */
export function renderFormPage(model: FormDesign, options: FormPageOptions): string {
  const size = formClientSize(model);
  const root = model.root;
  const caption = root ? controlCaption(root) || root.name : options.title;
  const stageWidth = twipsToPx(size.width);
  const stageHeight = twipsToPx(size.height);
  const body = renderFormBody(model, options.emptyText);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${options.nonce}';"> 
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(options.title)}</title>
<style nonce="${options.nonce}">
  html, body { margin: 0; padding: 0; }
  body {
    font-family: var(--vscode-font-family, "Segoe UI", Tahoma, sans-serif);
    color: var(--vscode-foreground, #333);
    background: var(--vscode-editor-background, #ffffff);
    padding: 16px;
  }
  .meta { display: flex; align-items: baseline; gap: 12px; margin-bottom: 12px; font-size: 12px; }
  .meta .title { font-size: 13px; font-weight: 600; }
  .meta .hint { opacity: .65; }
  .window { display: inline-block; max-width: 100%; border: 1px solid #b8b8b8; border-radius: 6px; overflow: hidden; box-shadow: 0 10px 28px rgba(0,0,0,.22); background: #f0f0f0; }
  .titlebar { height: 28px; display: flex; align-items: center; gap: 8px; padding: 0 10px; background: linear-gradient(#fdfdfd, #e8e8e8); border-bottom: 1px solid #c8c8c8; font-size: 12px; color: #222; }
  .titlebar .dot { width: 9px; height: 9px; border-radius: 50%; background: #4a90d9; flex: none; }
  .titlebar .caption { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .titlebar .close { width: 16px; height: 16px; display: grid; place-items: center; color: #666; font-size: 13px; }
  .stage { position: relative; overflow: hidden; background: #f0f0f0; }
  .stage.scroll { overflow: auto; }
  .control { position: absolute; box-sizing: border-box; }
  .vba-button { display: flex; align-items: center; justify-content: center; border: 1px solid #7a7a7a; border-radius: 2px; box-shadow: inset 0 1px 0 rgba(255,255,255,.8); overflow: hidden; padding: 0 6px; }
  .vba-button.pressed { box-shadow: inset 0 1px 3px rgba(0,0,0,.25); }
  .vba-button.disabled { color: #8a8a8a; }
  .vba-label { overflow: hidden; white-space: pre-wrap; }
  .vba-textbox { display: flex; align-items: center; border: 1px solid #7a7a7a; padding: 0 5px; overflow: hidden; white-space: pre; }
  .vba-combobox { display: flex; align-items: center; border: 1px solid #7a7a7a; padding-left: 5px; overflow: hidden; }
  .vba-combobox .arrow { flex: none; width: 18px; text-align: center; border-left: 1px solid #a0a0a0; background: #ececec; height: 100%; display: flex; align-items: center; justify-content: center; }
  .vba-combobox .grow, .vba-check .grow, .vba-option .grow { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; flex: 1; }
  .vba-listbox { border: 1px solid #7a7a7a; padding: 3px 5px; overflow: hidden; }
  .vba-listbox .row { height: 1.5em; border-bottom: 1px solid rgba(0,0,0,.05); }
  .vba-check, .vba-option { display: flex; align-items: center; gap: 5px; overflow: hidden; }
  .vba-check .mark, .vba-option .mark { width: 13px; height: 13px; display: inline-grid; place-items: center; flex: none; border: 1px solid #7a7a7a; background: #fff; font-size: 11px; line-height: 1; }
  .vba-option .mark { border-radius: 50%; }
  .vba-scroll { border: 1px solid #a8a8a8; background: #e9e9e9; display: flex; overflow: hidden; }
  .vba-scroll .thumb { background: #c5c5c5; border: 1px solid #a8a8a8; box-sizing: border-box; }
  .vba-scroll.vertical .thumb { width: 100%; height: 40%; }
  .vba-scroll.horizontal .thumb { height: 100%; width: 40%; }
  .vba-spin { display: flex; flex-direction: column; border: 1px solid #7a7a7a; }
  .vba-spin > div { flex: 1; display: grid; place-items: center; font-size: 8px; background: #f2f2f2; }
  .vba-spin .up { border-bottom: 1px solid #b0b0b0; }
  .vba-image { display: grid; place-items: center; color: #777; border: 1px solid #9a9a9a; background-image: linear-gradient(45deg, #d6d6d6 25%, transparent 25%, transparent 75%, #d6d6d6 75%), linear-gradient(45deg, #d6d6d6 25%, transparent 25%, transparent 75%, #d6d6d6 75%); background-size: 12px 12px; background-position: 0 0, 6px 6px; }
  .vba-line { min-height: 1px; }
  .vba-container { border: 1px solid #a0a0a0; overflow: visible; }
  .vba-frame-caption, .vba-tab { position: absolute; left: 6px; top: -8px; padding: 0 3px; font-size: 11px; background: #f0f0f0; color: #333; white-space: nowrap; }
  .vba-container.tabs { border-top: none; }
  .vba-container.tabs .vba-tab { top: 0; left: 0; height: 20px; display: flex; align-items: center; padding: 0 10px; border: 1px solid #a0a0a0; border-bottom: none; background: #f7f7f7; }
  .vba-container .vba-client { position: absolute; inset: 0; overflow: hidden; }
  .vba-container.tabs .vba-client { top: 20px; }
  .vba-unknown { display: grid; place-items: center; color: #888; font-size: 11px; }
  .empty { opacity: .7; font-size: 13px; }
  .error { color: var(--vscode-errorForeground, #c33); }
</style>
</head>
<body>
  <div class="meta"><span class="title" id="banner">${escapeHtml(options.title)}</span><span class="hint" id="hint">${escapeHtml(options.hint)}</span></div>
  <div class="window" id="window">
    <div class="titlebar">
      <span class="dot"></span>
      <span class="caption" id="form-caption">${escapeHtml(caption)}</span>
      <span class="close">✕</span>
    </div>
    <div class="stage" id="stage" style="width:${stageWidth}px;height:${stageHeight}px">${body}</div>
  </div>
<script nonce="${options.nonce}">
  (function () {
    const stage = document.getElementById('stage');
    const banner = document.getElementById('banner');
    const caption = document.getElementById('form-caption');
    window.addEventListener('message', function (event) {
      const data = event.data;
      if (!data || data.type !== 'render') { return; }
      stage.innerHTML = data.html;
      stage.style.width = data.width + 'px';
      stage.style.height = data.height + 'px';
      banner.textContent = data.title;
      caption.textContent = data.caption || data.title;
    });
  })();
</script>
</body>
</html>`;
}

/** 供热更新使用的数据：stage 内部 HTML 与尺寸。 */
export function buildRenderPayload(model: FormDesign, title: string, emptyText: string): { html: string; width: number; height: number; title: string; caption: string } {
  const size = formClientSize(model);
  const root = model.root;
  return {
    html: renderFormBody(model, emptyText),
    width: twipsToPx(size.width),
    height: twipsToPx(size.height),
    title,
    caption: root ? controlCaption(root) || root.name : '',
  };
}

// 保持导出面稳定，便于测试直接断言属性解析工具。
export { parsePropertyValue, parseNumericProperty };
