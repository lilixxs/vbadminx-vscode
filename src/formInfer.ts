/**
 * 为**没有设计段**的 .frm 文件提供兜底预览模型。
 *
 * 有些导出工具（例如只同步代码的 VBA 同步工具）写出的 .frm 只含代码，
 * `VERSION/Begin ... End` 设计块留在 Excel 文件里。本模块扫描代码中的
 * `With Me` / `With .控件` 与属性赋值，推断窗体尺寸、标题、控件类型与位置；
 * 信息不足时按控件名前缀与事件过程名补全，并自动纵向排列。
 *
 * 这是**近似推断**，调用方必须在预览中明确提示，不能当作设计器的真实布局。
 */
import {
  parseNumericProperty,
  parsePropertyValue,
  type FormControl,
  type FormDesign,
} from './formDesign';

const POINTS_TO_TWIPS = 20;

const FONT_ALIASES: Record<string, string> = {
  'font.name': 'FontName',
  'font.size': 'FontSize',
  'font.bold': 'FontBold',
  'font.italic': 'FontItalic',
  'font.underline': 'FontUnderline',
  'font.strikethrough': 'FontStrikethrough',
};

const FORM_PROPERTIES = new Set([
  'Width', 'Height', 'Left', 'Top', 'ClientWidth', 'ClientHeight', 'Caption',
  'BackColor', 'ForeColor', 'StartUpPosition', 'BorderStyle', 'SpecialEffect',
  'Font', 'FontName', 'FontSize', 'FontBold', 'FontItalic',
]);

const POSITION_PROPERTIES = new Set(['Width', 'Height', 'Left', 'Top', 'ClientWidth', 'ClientHeight']);

const PREFIX_TYPES: Array<[RegExp, string]> = [
  [/^(btn|cmd|command)/i, 'CommandButton'],
  [/^(lbl|label)/i, 'Label'],
  [/^(txt|tb|textbox)/i, 'TextBox'],
  [/^(cmb|cbo|combobox)/i, 'ComboBox'],
  [/^(lst|listbox)/i, 'ListBox'],
  [/^(chk|checkbox)/i, 'CheckBox'],
  [/^(opt|radio|option)/i, 'OptionButton'],
  [/^(tog|toggle)/i, 'ToggleButton'],
  [/^(fra|frame)/i, 'Frame'],
  [/^(img|image|pic)/i, 'Image'],
  [/^(spn|spin)/i, 'SpinButton'],
  [/^(sbr|scr|scroll)/i, 'ScrollBar'],
  [/^(tab|tabstrip)/i, 'TabStrip'],
  [/^(mpg|multipage)/i, 'MultiPage'],
];

/** 惯例默认尺寸（twips）：1 pt = 20 twips，按 96 DPI 近似。 */
const DEFAULT_SIZE: Record<string, { width: number; height: number }> = {
  CommandButton: { width: 1800, height: 480 },
  Label: { width: 2400, height: 320 },
  TextBox: { width: 2400, height: 400 },
  ComboBox: { width: 2400, height: 400 },
  ListBox: { width: 2400, height: 1200 },
  CheckBox: { width: 2000, height: 360 },
  OptionButton: { width: 2000, height: 360 },
  ToggleButton: { width: 1800, height: 480 },
  Frame: { width: 3000, height: 1600 },
  Image: { width: 2400, height: 1600 },
  ScrollBar: { width: 240, height: 1200 },
  SpinButton: { width: 240, height: 480 },
};

interface MutableControl {
  name: string;
  type: string;
  properties: Record<string, string>;
  /** 是否已有显式位置（Left/Top）；用于只排列缺失位置的控件。 */
  positioned: boolean;
}

function guessType(name: string, event?: string): string {
  for (const [pattern, type] of PREFIX_TYPES) {
    if (pattern.test(name)) { return type; }
  }
  if (event === 'Click' || /^(ok|cancel|close|save|add|delete|del|edit|update|submit|confirm|run|exit|back|next|prev)/i.test(name)) {
    return 'CommandButton';
  }
  return 'Label';
}

/** 找到引号外的等号。 */
function findEquals(line: string): number {
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]!;
    if (char === '"') { quoted = !quoted; continue; }
    if (!quoted && char === '=') { return index; }
  }
  return -1;
}

function currentOwner(stack: string[]): string | undefined {
  return stack.length ? stack[stack.length - 1] : undefined;
}

export function inferFormDesignFromCode(text: string, formName = 'UserForm'): FormDesign {
  const formProperties: Record<string, string> = {};
  const controls = new Map<string, MutableControl>();
  const withStack: string[] = [];
  const header = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/im.exec(text)?.[1];

  const controlOf = (name: string, event?: string): MutableControl => {
    let control = controls.get(name);
    if (!control) {
      control = { name, type: guessType(name, event), properties: {}, positioned: false };
      controls.set(name, control);
    } else if (event && control.type === 'Label' && guessType(name, event) === 'CommandButton') {
      control.type = 'CommandButton';
    }
    return control;
  };

  const assign = (owner: string | undefined, path: string, rawValue: string): void => {
    const key = FONT_ALIASES[path.toLowerCase()] ?? path;
    if (!key) { return; }
    const value = parsePropertyValue(rawValue);
    let stored = value;
    if (POSITION_PROPERTIES.has(key)) {
      const number = parseNumericProperty(value, Number.NaN);
      if (Number.isFinite(number)) { stored = String(Math.round(number * POINTS_TO_TWIPS)); }
    }
    if (owner === undefined) {
      formProperties[key] = stored;
      return;
    }
    const control = controlOf(owner);
    control.properties[key] = stored;
    if (key === 'Left' || key === 'Top') { control.positioned = true; }
    if (key === 'Width' || key === 'Height') { control.positioned = control.positioned || control.properties['Left'] !== undefined || control.properties['Top'] !== undefined; }
  };

  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("'")) { continue; }

    const withMatch = /^With\s+(.+?)\s*(?:'.*)?$/i.exec(trimmed);
    if (withMatch) {
      let expression = withMatch[1]!.trim();
      expression = expression.replace(/\s+'.*$/, '').trim();
      let owner: string | undefined;
      if (/^Me$/i.test(expression)) { owner = undefined; }
      else if (/^Me\.([A-Za-z_]\w*)$/i.test(expression)) { owner = /^Me\.([A-Za-z_]\w*)$/i.exec(expression)![1]!; }
      else if (/^\.[A-Za-z_]\w*$/.test(expression)) { owner = expression.slice(1); }
      else if (/^[A-Za-z_]\w*$/.test(expression) && !FORM_PROPERTIES.has(expression)) { owner = expression; }
      withStack.push(owner ?? '');
      if (owner) { controlOf(owner); }
      continue;
    }
    if (/^End\s+With\b/i.test(trimmed)) { withStack.pop(); continue; }

    const eventMatch = /^\s*(?:Private|Public|Friend)?\s*Sub\s+([A-Za-z_]\w*)_([A-Za-z]+)\s*\(/i.exec(line);
    if (eventMatch && !/^UserForm$/i.test(eventMatch[1]!)) {
      controlOf(eventMatch[1]!, eventMatch[2]);
      continue;
    }

    const equals = findEquals(line);
    if (equals < 0) { continue; }
    let left = line.slice(0, equals).trim().replace(/^\./, '');
    const right = line.slice(equals + 1).trim();
    if (!left || !right || !/^[A-Za-z_]\w*(\.[A-Za-z_]\w*)*$/.test(left)) { continue; }
    const parts = left.split('.');
    if (parts[0]!.toLowerCase() === 'me') { parts.shift(); }
    if (!parts.length) { continue; }
    const first = parts[0]!;
    if (FORM_PROPERTIES.has(first)) {
      const owner = currentOwner(withStack);
      assign(owner === '' ? undefined : owner, parts.join('.'), right);
      continue;
    }
    const owner = first;
    const path = parts.slice(1).join('.');
    if (!path) { continue; }
    assign(owner, path, right);
  }

  const children = [...controls.values()];
  if (!children.length && !Object.keys(formProperties).length) { return { root: undefined }; }

  const root: FormControl = {
    name: header ?? formName,
    type: 'UserForm',
    rawType: 'UserForm',
    properties: { ...formProperties },
    children: [],
  };
  if (!root.properties['Caption']) { root.properties['Caption'] = header ?? formName; }

  let width = parseNumericProperty(root.properties['ClientWidth'] ?? root.properties['Width'], 0);
  let height = parseNumericProperty(root.properties['ClientHeight'] ?? root.properties['Height'], 0);
  if (width <= 0) { width = 6000; }
  if (height <= 0) { height = 4000; }

  // 只排列缺失位置的控件；有显式坐标的保持代码中的布局。
  let cursor = 240;
  const margin = 240;
  const gap = 120;
  for (const control of children) {
    const defaults = DEFAULT_SIZE[control.type] ?? { width: 2400, height: 480 };
    if (control.properties['Left'] === undefined) { control.properties['Left'] = String(margin); }
    if (control.properties['Top'] === undefined) {
      control.properties['Top'] = String(cursor);
      cursor += parseNumericProperty(control.properties['Height'] ?? String(defaults.height), defaults.height) + gap;
    }
    if (control.properties['Width'] === undefined) { control.properties['Width'] = String(defaults.width); }
    if (control.properties['Height'] === undefined) { control.properties['Height'] = String(defaults.height); }
    if (!control.properties['Caption'] && (control.type === 'CommandButton' || control.type === 'Label' || control.type === 'CheckBox' || control.type === 'OptionButton' || control.type === 'ToggleButton')) {
      control.properties['Caption'] = control.name;
    }
  }

  // 按位置排序，保证渲染层级接近视觉顺序。
  children.sort((a, b) => parseNumericProperty(a.properties['Top'], 0) - parseNumericProperty(b.properties['Top'], 0)
    || parseNumericProperty(a.properties['Left'], 0) - parseNumericProperty(b.properties['Left'], 0));

  for (const control of children) {
    const bottom = parseNumericProperty(control.properties['Top'], 0) + parseNumericProperty(control.properties['Height'], 0);
    if (bottom + gap > height) { height = bottom + gap; }
    control.properties['Visible'] ??= 'True';
    root.children.push({ name: control.name, type: control.type, rawType: control.type, properties: control.properties, children: [] });
  }

  root.properties['ClientWidth'] = String(width);
  root.properties['ClientHeight'] = String(height);
  return { root, inferred: true };
}
