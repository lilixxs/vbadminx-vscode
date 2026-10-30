/**
 * VBA 窗体（.frm）设计段解析（不依赖 vscode 模块）。
 *
 * 只读取文本形式的设计信息：
 *   Begin {GUID} UserForm1
 *      Caption = "..."
 *      Begin VB.CommandButton Command1
 *         ...
 *      End
 *   End
 *
 * `.frx`（二进制资源）不解析，图片类控件以占位框呈现。
 */

export interface FormControl {
  /** 控件名（Begin 行最后一个记号）。 */
  name: string;
  /** 归一化后的控件类型，例如 CommandButton；未知 GUID 保留原始文本。 */
  type: string;
  /** Begin 行中的原始类型文本（GUID 或 VB.Xxx）。 */
  rawType: string;
  /** 属性名到值的映射；BeginProperty Font 内的属性记为 Font.Name 等。 */
  properties: Record<string, string>;
  children: FormControl[];
}

export interface FormDesign {
  root?: FormControl;
}

/** 常见 MSForms 控件 GUID（大写）。未列出的 GUID 会显示为未知控件。 */
const GUID_TYPES: Record<string, string> = {
  'C62A69F0-16DC-11CE-9E98-00AA00574A4F': 'UserForm',
  'D7053240-CE69-11CD-A777-00DD01143C57': 'CommandButton',
  '978C9E23-D4B0-11CE-BF2D-00AA003F40D0': 'Label',
  '8BD21D10-EC42-11CE-9E0D-00AA006002F3': 'TextBox',
  '8BD21D20-EC42-11CE-9E0D-00AA006002F3': 'ListBox',
  '8BD21D30-EC42-11CE-9E0D-00AA006002F3': 'ComboBox',
  '8BD21D40-EC42-11CE-9E0D-00AA006002F3': 'CheckBox',
  '8BD21D50-EC42-11CE-9E0D-00AA006002F3': 'OptionButton',
  '8BD21D60-EC42-11CE-9E0D-00AA006002F3': 'ToggleButton',
  '6E182020-F460-11CE-9BCD-00AA00608E01': 'Frame',
  '4C599241-6926-101B-9992-00000B65C6F9': 'Image',
  'DFD181E0-5E2F-11CE-A449-00AA004A803D': 'ScrollBar',
  '79176FB0-B7F2-11CE-97EF-00AA006D2776': 'SpinButton',
  'EAE50EB0-4A62-11CE-BED6-00AA00611080': 'TabStrip',
  '46E31370-3F7A-11CE-BED6-00AA00611080': 'MultiPage',
};

/** 把 Begin 行中的类型文本归一化为短类型名。 */
export function resolveControlType(rawType: string): string {
  const trimmed = rawType.trim();
  const guid = /^\{([0-9A-Fa-f-]+)\}$/.exec(trimmed);
  if (guid) {
    const key = guid[1]!.toUpperCase();
    return GUID_TYPES[key] ?? `Unknown (${key})`;
  }
  const last = trimmed.split('.').pop() ?? trimmed;
  return last || 'Unknown';
}

/** 取出属性行右侧的值：支持带双引号的字符串与行尾注释。 */
export function parsePropertyValue(raw: string): string {
  const text = raw.trim();
  if (!text) { return ''; }
  if (text.startsWith('"')) {
    let value = '';
    let index = 1;
    while (index < text.length) {
      const char = text[index]!;
      if (char === '"') {
        if (text[index + 1] === '"') { value += '"'; index += 2; continue; }
        break;
      }
      value += char;
      index += 1;
    }
    return value;
  }
  const comment = text.indexOf("'");
  return (comment >= 0 ? text.slice(0, comment) : text).trim();
}

/** 解析数字属性：支持十进制与 &H 十六进制（忽略末尾的 &）。 */
export function parseNumericProperty(value: string | undefined, fallback = 0): number {
  if (value === undefined) { return fallback; }
  const text = value.trim();
  if (!text) { return fallback; }
  const hex = /^&H([0-9A-Fa-f]+)&?$/.exec(text);
  if (hex) { return Number.parseInt(hex[1]!, 16); }
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : fallback;
}

const SYSTEM_COLORS: Record<number, string> = {
  2: '#99b4d1', // COLOR_ACTIVECAPTION
  5: '#ffffff', // COLOR_WINDOW
  8: '#000000', // COLOR_WINDOWTEXT
  9: '#6d6d6d', // COLOR_GRAYTEXT
  13: '#0078d7', // COLOR_HIGHLIGHT
  14: '#ffffff', // COLOR_HIGHLIGHTTEXT
  15: '#f0f0f0', // COLOR_BTNFACE
  16: '#a0a0a0', // COLOR_BTNSHADOW
  18: '#000000', // COLOR_BTNTEXT
};

/** OLE 颜色值转 CSS：0x00BBGGRR，或 0x80000000 | 系统颜色索引。 */
export function oleColorToCss(value: string | undefined, fallback: string): string {
  if (value === undefined) { return fallback; }
  const parsed = parseNumericProperty(value, Number.NaN);
  if (!Number.isFinite(parsed)) { return fallback; }
  const unsigned = parsed >>> 0;
  if (((unsigned & 0xff000000) >>> 0) === 0x80000000) {
    return SYSTEM_COLORS[unsigned & 0xff] ?? fallback;
  }
  const red = unsigned & 0xff;
  const green = (unsigned >>> 8) & 0xff;
  const blue = (unsigned >>> 16) & 0xff;
  return `rgb(${red}, ${green}, ${blue})`;
}

export function parseFormDesign(text: string): FormDesign {
  const stack: FormControl[] = [];
  const propertyStack: string[] = [];
  let root: FormControl | undefined;
  for (const line of text.split(/\r?\n/)) {
    if (/^\s*Begin\b/i.test(line) && !/^\s*BeginProperty\b/i.test(line)) {
      const rest = line.replace(/^\s*Begin\s+/i, '').trim();
      const match = /^(\S+)\s+(\S+)\s*$/.exec(rest);
      if (!match) { continue; }
      const rawType = match[1]!;
      const control: FormControl = {
        name: match[2]!,
        type: resolveControlType(rawType),
        rawType,
        properties: {},
        children: [],
      };
      const parent = stack[stack.length - 1];
      if (parent) { parent.children.push(control); } else { root ??= control; }
      stack.push(control);
      continue;
    }
    if (/^\s*BeginProperty\b/i.test(line)) {
      const match = /^\s*BeginProperty\s+(\w+)/i.exec(line);
      propertyStack.push(match?.[1] ?? 'Property');
      continue;
    }
    if (/^\s*EndProperty\s*$/i.test(line)) {
      propertyStack.pop();
      continue;
    }
    if (/^\s*End\s*$/i.test(line)) {
      stack.pop();
      continue;
    }
    const current = stack[stack.length - 1];
    if (!current) { continue; }
    const match = /^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.*)$/.exec(line);
    if (!match) { continue; }
    const key = propertyStack.length ? `${propertyStack[propertyStack.length - 1]}.${match[1]!}` : match[1]!;
    current.properties[key] = parsePropertyValue(match[2]!);
  }
  return { root };
}

function property(control: FormControl, name: string): string | undefined {
  return control.properties[name];
}

export function controlCaption(control: FormControl): string {
  const raw = property(control, 'Caption') ?? '';
  // “&” 表示访问键：&& 是字面量 &，单个 & 后一个字符加下划线（这里只做去标记处理）。
  return raw.replaceAll('&&', '\u0000').replace(/&/g, '').replaceAll('\u0000', '&');
}

export function controlFontSize(control: FormControl, fallback = 9): number {
  return parseNumericProperty(property(control, 'FontSize') ?? property(control, 'Font.Size'), fallback);
}

export function controlFontName(control: FormControl): string {
  return property(control, 'FontName') ?? property(control, 'Font.Name') ?? 'Tahoma';
}

export function controlBold(control: FormControl): boolean {
  const bold = property(control, 'FontBold');
  if (bold !== undefined) { return bold === 'True' || bold === '-1'; }
  const weight = parseNumericProperty(property(control, 'Font.Weight'), 400);
  return weight >= 700;
}

export function controlItalic(control: FormControl): boolean {
  return property(control, 'FontItalic') === 'True'
    || property(control, 'Font.Italic') === 'True';
}

export function controlVisible(control: FormControl): boolean {
  const value = property(control, 'Visible');
  return value === undefined || value === 'True' || value === '-1';
}

/** 已启用状态：0/False 表示禁用（预览里只改样式）。 */
export function controlEnabled(control: FormControl): boolean {
  const value = property(control, 'Enabled');
  return value === undefined || value === 'True' || value === '-1';
}
