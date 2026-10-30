/**
 * VBA 源码格式化。
 *
 * 需求依据 `docs/requirements.md`：
 * - FMT-01 默认 4 空格缩进，可配置；
 * - FMT-02 关键字按 VBA 标准大小写，可配置关闭；
 * - FMT-03 默认不折行、不重排声明、不格式化设计段；
 * - FMT-04 不改变注释内容、字符串内容与条件编译语义。
 *
 * 本模块不依赖 VBA 解析器：只做“物理行 + 逻辑行”级扫描。
 * 遇到无法确定的输入（未闭合字符串、Attribute、窗体设计段）时保持原样，
 * 宁可少改，也不猜测。
 */

export interface FormatOptions {
  tabSize?: number;
  insertSpaces?: boolean;
  keywordCase?: boolean;
  lineWrap?: boolean;
  reorderDeclarations?: boolean;
  formatDesign?: boolean;
}

const DEFAULT_TAB_SIZE = 4;
const LINE_WRAP_LENGTH = 120;
const MASKED = '\u0000';

interface SourceLine {
  text: string;
  eol: string;
}

interface LineMask {
  /** 与 text 等长。代码字符保留原样，字符串/注释/日期字面量/方括号标识符替换为 \u0000。 */
  mask: string;
  hasComment: boolean;
  commentStart: number;
  startsInString: boolean;
  endsInString: boolean;
}

type LineKind = 'code' | 'comment' | 'blank' | 'directive' | 'metadata' | 'frozen';

interface LinePlan {
  kind: LineKind;
  /** 该行应使用的缩进级别。 */
  level: number;
  /** 该行是否是逻辑语句的起始物理行。 */
  statementStart: boolean;
}

interface MetadataInfo {
  isMetadata: boolean[];
  /** -1 表示不属于设计段；否则为该行在设计段内的嵌套深度。 */
  designLevel: number[];
}

type FrameKind = 'proc' | 'if' | 'select' | 'case' | 'for' | 'do' | 'while' | 'with' | 'type' | 'enum';

interface ResolvedOptions {
  unit: string;
  keywordCase: boolean;
  lineWrap: boolean;
  reorder: boolean;
  formatDesign: boolean;
}

/* ------------------------------------------------------------------ */
/* 入口                                                                */
/* ------------------------------------------------------------------ */

/**
 * 格式化 VBA 源码。
 *
 * @param text 完整文件内容，可含 BOM 与 CRLF/LF 混排。
 * @param options 可选格式化选项，默认 4 空格缩进、关键字标准化。
 * @param lineRange 0 基、闭区间的行范围（与 LSP `Range` 一致）。省略时格式化全文。
 * @returns 新文本；BOM 与每行原换行形式保持不变。
 */
export function formatVba(
  text: string,
  options: FormatOptions = {},
  lineRange?: { start: number; end: number },
): string {
  const settings = resolveOptions(options);
  const bom = text.startsWith('\uFEFF') ? '\uFEFF' : '';
  const body = bom === '' ? text : text.slice(1);
  const lines = splitSource(body);

  const masks = buildMasks(lines);
  const meta = classifyMetadata(lines, masks);
  const plans = planLevels(lines, masks, meta);

  const last = lines.length - 1;
  let start = 0;
  let end = last;
  if (lineRange !== undefined) {
    if (lineRange.start > lineRange.end) return text;
    start = clamp(lineRange.start, 0, last);
    end = clamp(lineRange.end, 0, last);
    if (end < start) return text;
  }

  const output: SourceLine[] = lines.map((line) => ({ text: line.text, eol: line.eol }));

  for (let i = start; i <= end; i += 1) {
    const plan = plans[i];
    const mask = masks[i];
    const target = output[i];
    if (plan === undefined || mask === undefined || target === undefined) continue;

    if (plan.kind === 'metadata') {
      if (!settings.formatDesign) continue;
      const stripped = target.text.replace(/^[ \t]+/, '');
      if (stripped === '') continue;
      if (/^attribute\b/i.test(stripped)) {
        target.text = stripped;
        continue;
      }
      const level = Math.max(meta.designLevel[i] ?? 0, 0);
      target.text = settings.unit.repeat(level) + stripped;
      continue;
    }
    if (plan.kind === 'frozen') continue;
    if (plan.kind === 'blank') {
      target.text = '';
      continue;
    }

    const content = settings.keywordCase ? caseKeywords(target.text, mask.mask) : target.text;
    const stripped = content.replace(/^[ \t]+/, '');
    target.text = stripped === '' ? '' : settings.unit.repeat(plan.level) + stripped;
  }

  if (settings.reorder) reorderDimRuns(output, start, end);

  const finalLines = settings.lineWrap ? applyLineWrap(output, plans, settings, start, end) : output;
  return bom + finalLines.map((line) => line.text + line.eol).join('');
}

function resolveOptions(options: FormatOptions): ResolvedOptions {
  const requested = options.tabSize;
  const tabSize =
    typeof requested === 'number' && Number.isFinite(requested) && requested > 0
      ? Math.floor(requested)
      : DEFAULT_TAB_SIZE;
  const insertSpaces = options.insertSpaces !== false;
  return {
    unit: insertSpaces ? ' '.repeat(tabSize) : '\t',
    keywordCase: options.keywordCase !== false,
    lineWrap: options.lineWrap === true,
    reorder: options.reorderDeclarations === true,
    formatDesign: options.formatDesign === true,
  };
}

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.min(Math.max(Math.floor(value), min), max);
}

/* ------------------------------------------------------------------ */
/* 物理行切分                                                          */
/* ------------------------------------------------------------------ */

function splitSource(body: string): SourceLine[] {
  const lines: SourceLine[] = [];
  let start = 0;
  for (let i = 0; i < body.length; i += 1) {
    const ch = body.charAt(i);
    if (ch !== '\r' && ch !== '\n') continue;
    let eol = ch;
    if (ch === '\r' && body.charAt(i + 1) === '\n') {
      eol = '\r\n';
      i += 1;
    }
    lines.push({ text: body.slice(start, i + 1 - eol.length), eol });
    start = i + 1;
  }
  lines.push({ text: body.slice(start), eol: '' });
  return lines;
}

/* ------------------------------------------------------------------ */
/* 掩码：区分代码、字符串、注释、日期字面量                            */
/* ------------------------------------------------------------------ */

const DATE_LITERAL = /^#(?:\s*\d{1,4}[\/\-:]\d[^#]*|\s*\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}[^#]*)#/;

function buildMasks(lines: SourceLine[]): LineMask[] {
  const result: LineMask[] = [];
  let carry = false;
  for (const line of lines) {
    const mask = maskLine(line.text, carry);
    result.push(mask);
    carry = mask.endsInString;
  }
  return result;
}

function maskLine(text: string, startsInString: boolean): LineMask {
  const out = new Array<string>(text.length).fill(MASKED);
  let hasComment = false;
  let commentStart = -1;
  let endsInString = false;
  let i = 0;

  if (startsInString) {
    let closed = false;
    while (i < text.length) {
      if (text.charAt(i) === '"') {
        if (text.charAt(i + 1) === '"') {
          i += 2;
          continue;
        }
        i += 1;
        closed = true;
        break;
      }
      i += 1;
    }
    if (!closed) {
      return { mask: out.join(''), hasComment, commentStart, startsInString, endsInString: true };
    }
  }

  while (i < text.length) {
    const ch = text.charAt(i);
    if (ch === '"') {
      let j = i + 1;
      let closed = false;
      while (j < text.length) {
        if (text.charAt(j) === '"') {
          if (text.charAt(j + 1) === '"') {
            j += 2;
            continue;
          }
          j += 1;
          closed = true;
          break;
        }
        j += 1;
      }
      if (!closed) {
        endsInString = true;
        i = text.length;
        break;
      }
      i = j;
      continue;
    }
    if (ch === "'") {
      hasComment = true;
      commentStart = i;
      i = text.length;
      break;
    }
    if (ch === '[') {
      const close = text.indexOf(']', i + 1);
      if (close >= 0) {
        i = close + 1;
        continue;
      }
      out[i] = ch;
      i += 1;
      continue;
    }
    if ((ch === 'r' || ch === 'R') && isRemCommentStart(out, i) && /^rem(?:\s|$)/i.test(text.slice(i, i + 4))) {
      hasComment = true;
      commentStart = i;
      i = text.length;
      break;
    }
    if (ch === '#' && !precededByIdentifier(out, i) && matchDateLiteral(text, i) >= 0) {
      i = matchDateLiteral(text, i);
      continue;
    }
    out[i] = ch;
    i += 1;
  }

  return { mask: out.join(''), hasComment, commentStart, startsInString, endsInString };
}

function matchDateLiteral(text: string, start: number): number {
  const match = DATE_LITERAL.exec(text.slice(start));
  return match === null ? -1 : start + match[0].length;
}

function precededByIdentifier(out: string[], index: number): boolean {
  const prev = index - 1;
  if (prev < 0) return false;
  const ch = out[prev];
  if (ch === undefined || ch === MASKED) return false;
  return /[A-Za-z0-9_)]/.test(ch);
}

function isRemCommentStart(out: string[], index: number): boolean {
  for (let i = index - 1; i >= 0; i -= 1) {
    const ch = out[i];
    if (ch === undefined) continue;
    if (ch === ' ' || ch === '\t') continue;
    if (ch === MASKED) return false;
    return ch === ':';
  }
  return true;
}

function toCode(mask: string): string {
  return mask.split(MASKED).join(' ');
}

/* ------------------------------------------------------------------ */
/* Attribute / VBE 头 / 窗体设计段                                     */
/* ------------------------------------------------------------------ */

function classifyMetadata(lines: SourceLine[], masks: LineMask[]): MetadataInfo {
  const count = lines.length;
  const isMetadata = new Array<boolean>(count).fill(false);
  const designLevel = new Array<number>(count).fill(-1);

  let index = 0;
  let beginDepth = 0;
  let started = false;

  while (index < count) {
    const mask = masks[index];
    const code = mask === undefined ? '' : toCode(mask.mask).trim();
    if (code === '') {
      if (started) {
        isMetadata[index] = true;
        designLevel[index] = beginDepth;
      }
      index += 1;
      continue;
    }
    if (beginDepth > 0) {
      isMetadata[index] = true;
      if (/^begin\b/i.test(code)) {
        designLevel[index] = beginDepth;
        beginDepth += 1;
      } else if (/^end$/i.test(code)) {
        beginDepth -= 1;
        designLevel[index] = beginDepth;
      } else {
        designLevel[index] = beginDepth;
      }
      index += 1;
      continue;
    }
    if (/^version\b/i.test(code)) {
      started = true;
      isMetadata[index] = true;
      designLevel[index] = 0;
      index += 1;
      continue;
    }
    if (/^begin\b/i.test(code)) {
      started = true;
      isMetadata[index] = true;
      designLevel[index] = 0;
      beginDepth = 1;
      index += 1;
      continue;
    }
    if (/^attribute\b/i.test(code)) {
      isMetadata[index] = true;
      designLevel[index] = 0;
      index += 1;
      continue;
    }
    if (/^(?:multiuse|object|recalcengine|smallbusiness|lookingat)\b/i.test(code)) {
      isMetadata[index] = true;
      designLevel[index] = 0;
      index += 1;
      continue;
    }
    break;
  }

  for (let i = 0; i < count; i += 1) {
    const mask = masks[i];
    if (mask === undefined) continue;
    if (/^attribute\b/i.test(toCode(mask.mask).trim())) {
      isMetadata[i] = true;
      if ((designLevel[i] ?? -1) < 0) designLevel[i] = 0;
    }
  }

  return { isMetadata, designLevel };
}

/* ------------------------------------------------------------------ */
/* 缩进规划                                                            */
/* ------------------------------------------------------------------ */

function planLevels(lines: SourceLine[], masks: LineMask[], meta: MetadataInfo): LinePlan[] {
  const plans = new Array<LinePlan | undefined>(lines.length);
  const frames: FrameKind[] = [];
  // 条件编译指令不改变 VBA 块结构，只在排版上形成一层缩进。
  let directiveDepth = 0;
  let i = 0;

  while (i < lines.length) {
    const mask = masks[i];
    if (mask === undefined) {
      i += 1;
      continue;
    }
    if (meta.isMetadata[i] === true) {
      plans[i] = { kind: 'metadata', level: 0, statementStart: false };
      i += 1;
      continue;
    }
    if (mask.startsInString || mask.endsInString) {
      plans[i] = { kind: 'frozen', level: frames.length + directiveDepth, statementStart: false };
      i += 1;
      continue;
    }

    const code = toCode(mask.mask);
    const trimmed = code.trim();
    if (trimmed === '') {
      plans[i] = {
        kind: mask.hasComment ? 'comment' : 'blank',
        level: frames.length + directiveDepth,
        statementStart: false,
      };
      i += 1;
      continue;
    }
    if (trimmed.startsWith('#')) {
      const directive = /^#\s*(.*)$/.exec(trimmed)?.[1] ?? '';
      if (/^end\s+if\b/i.test(directive)) {
        directiveDepth = Math.max(directiveDepth - 1, 0);
        plans[i] = { kind: 'directive', level: frames.length + directiveDepth, statementStart: false };
      } else if (/^(?:elseif|else)\b/i.test(directive)) {
        plans[i] = {
          kind: 'directive',
          level: frames.length + Math.max(directiveDepth - 1, 0),
          statementStart: false,
        };
      } else {
        plans[i] = { kind: 'directive', level: frames.length + directiveDepth, statementStart: false };
        if (/^if\b/i.test(directive)) directiveDepth += 1;
      }
      i += 1;
      continue;
    }

    let end = i;
    while (end + 1 < lines.length) {
      const current = masks[end];
      if (current === undefined || !endsWithLineContinuation(current.mask)) break;
      const next = masks[end + 1];
      if (next === undefined) break;
      if (meta.isMetadata[end + 1] === true || next.startsInString || next.endsInString) break;
      end += 1;
    }

    const parts: string[] = [];
    for (let k = i; k <= end; k += 1) {
      const item = masks[k];
      if (item !== undefined) parts.push(toCode(item.mask));
    }
    const segments = splitSegments(parts.join(' '));
    const head = segments.find((segment) => segment.trim() !== '') ?? '';

    let level: number;
    if (segments.length > 1 && isLabelSegment(head)) {
      level = Math.max(frames.length - 1, 0);
    } else {
      level = analyzeSegments(frames, segments);
    }
    level += directiveDepth;

    plans[i] = { kind: 'code', level, statementStart: true };
    for (let k = i + 1; k <= end; k += 1) {
      plans[k] = { kind: 'code', level: level + 1, statementStart: false };
    }
    i = end + 1;
  }

  return plans.map((plan) => plan ?? { kind: 'blank', level: 0, statementStart: false });
}

function endsWithLineContinuation(mask: string): boolean {
  let j = mask.length - 1;
  while (j >= 0) {
    const ch = mask.charAt(j);
    if (ch === ' ' || ch === '\t' || ch === '\r' || ch === MASKED) {
      j -= 1;
      continue;
    }
    break;
  }
  if (j <= 0) return false;
  if (mask.charAt(j) !== '_') return false;
  const prev = mask.charAt(j - 1);
  return prev === ' ' || prev === '\t';
}

function splitSegments(code: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < code.length; i += 1) {
    const ch = code.charAt(i);
    if (ch === '(') depth += 1;
    else if (ch === ')') depth = Math.max(0, depth - 1);
    else if (ch === ':' && depth === 0 && code.charAt(i + 1) !== '=') {
      parts.push(code.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(code.slice(start));
  return parts;
}

const LABEL_FORBIDDEN =
  /^(?:else|elseif|case|end|next|loop|wend|do|for|if|select|with|type|enum|sub|function|property|dim|const|set|let|on|exit|goto|gosub|option|declare|attribute|rem|while|stop|randomize)\b/i;

function isLabelSegment(segment: string): boolean {
  const text = segment.trim();
  if (!/^(?:[A-Za-z_][A-Za-z0-9_]*|\d+)$/.test(text)) return false;
  return !LABEL_FORBIDDEN.test(text);
}

function analyzeSegments(frames: FrameKind[], segments: string[]): number {
  let level = frames.length;
  let first = true;
  for (const segment of segments) {
    const text = segment.trim();
    if (text === '') continue;
    const printed = processSegment(frames, text);
    if (first) {
      level = printed;
      first = false;
    }
  }
  return level;
}

function popFrame(frames: FrameKind[], kind: FrameKind): void {
  if (frames.length > 0 && frames[frames.length - 1] === kind) frames.pop();
}

function processSegment(frames: FrameKind[], text: string): number {
  const endMatch = /^end\s+(sub|function|property|if|with|select|type|enum)\b/i.exec(text);
  if (endMatch !== null) {
    const tail = (endMatch[1] ?? '').toLowerCase();
    if (tail === 'sub' || tail === 'function' || tail === 'property') popFrame(frames, 'proc');
    else if (tail === 'if') popFrame(frames, 'if');
    else if (tail === 'with') popFrame(frames, 'with');
    else if (tail === 'select') {
      popFrame(frames, 'case');
      popFrame(frames, 'select');
    } else if (tail === 'type') popFrame(frames, 'type');
    else if (tail === 'enum') popFrame(frames, 'enum');
    return frames.length;
  }

  if (/^next\b/i.test(text)) {
    const rest = /^next\s+(.*)$/i.exec(text);
    const names = (rest?.[1] ?? '').split(',').filter((name) => name.trim() !== '');
    const count = Math.max(1, names.length);
    for (let i = 0; i < count; i += 1) popFrame(frames, 'for');
    return frames.length;
  }

  if (/^loop\b/i.test(text)) {
    popFrame(frames, 'do');
    return frames.length;
  }

  if (/^wend\b/i.test(text)) {
    popFrame(frames, 'while');
    return frames.length;
  }

  if (/^(?:elseif|else)\b/i.test(text)) {
    return Math.max(frames.length - 1, 0);
  }

  if (/^case\b/i.test(text)) {
    popFrame(frames, 'case');
    const level = frames.length;
    frames.push('case');
    return level;
  }

  const level = frames.length;
  const opened = openingFrame(text);
  if (opened !== undefined) frames.push(opened);
  return level;
}

function openingFrame(text: string): FrameKind | undefined {
  if (/^(?:(?:public|private|friend|static)\s+)*declare\b/i.test(text)) return undefined;
  if (/^(?:(?:public|private|friend|static)\s+)*(?:sub|function|property)\b/i.test(text)) return 'proc';
  if (/^if\b/i.test(text) && /\bthen\s*$/i.test(text)) return 'if';
  if (/^for\b/i.test(text)) return 'for';
  if (/^do\b/i.test(text)) return 'do';
  if (/^while\b/i.test(text)) return 'while';
  if (/^with\b/i.test(text)) return 'with';
  if (/^select\s+case\b/i.test(text)) return 'select';
  if (/^(?:(?:public|private|friend)\s+)?type\b/i.test(text)) return 'type';
  if (/^(?:(?:public|private|friend)\s+)?enum\b/i.test(text)) return 'enum';
  return undefined;
}

/* ------------------------------------------------------------------ */
/* 关键字大小写                                                        */
/* ------------------------------------------------------------------ */

const KEYWORD_LIST = [
  'AddressOf', 'Alias', 'And', 'Any', 'As', 'Attribute',
  'Base', 'Binary', 'Boolean', 'ByRef', 'Byte', 'ByVal',
  'Call', 'Case', 'Close', 'Compare', 'Const', 'Currency',
  'Date', 'Decimal', 'Declare', 'Dim', 'Do', 'Double',
  'Each', 'Else', 'ElseIf', 'Empty', 'End', 'Enum', 'Eqv', 'Erase', 'Error', 'Event', 'Exit', 'Explicit',
  'False', 'For', 'Friend', 'Function',
  'Get', 'Global', 'GoSub', 'GoTo',
  'If', 'Implements', 'Imp', 'In', 'Input', 'Integer', 'Is',
  'Let', 'Lib', 'Like', 'Lock', 'Long', 'LongLong', 'LongPtr', 'Loop', 'LSet',
  'Me', 'Mod', 'New', 'Next', 'Not', 'Nothing', 'Null',
  'Object', 'On', 'Open', 'Option', 'Optional', 'Or', 'Output',
  'ParamArray', 'Preserve', 'Print', 'Private', 'Property', 'PtrSafe', 'Public', 'Put',
  'RaiseEvent', 'Random', 'ReDim', 'Resume', 'Return', 'RSet',
  'Seek', 'Select', 'Set', 'Shared', 'Single', 'Static', 'Step', 'Stop', 'String', 'Sub',
  'Text', 'Then', 'To', 'True', 'Type', 'TypeOf',
  'Unlock', 'Until', 'Variant',
  'Wend', 'While', 'With', 'WithEvents',
  'Xor',
];

const KEYWORDS = new Map<string, string>();
for (const word of KEYWORD_LIST) KEYWORDS.set(word.toLowerCase(), word);

function caseKeywords(text: string, mask: string): string {
  let result = text;
  let i = 0;
  while (i < text.length) {
    if (mask.charAt(i) === MASKED) {
      i += 1;
      continue;
    }
    const ch = text.charAt(i);
    if (!/[A-Za-z_]/.test(ch)) {
      i += 1;
      continue;
    }
    let j = i + 1;
    while (j < text.length && mask.charAt(j) !== MASKED && /[A-Za-z0-9_]/.test(text.charAt(j))) j += 1;
    const word = text.slice(i, j);
    const canonical = KEYWORDS.get(word.toLowerCase());
    if (canonical !== undefined && canonical !== word && previousCodeChar(text, mask, i) !== '.') {
      result = result.slice(0, i) + canonical + result.slice(j);
    }
    i = j;
  }
  return result;
}

function previousCodeChar(text: string, mask: string, index: number): string {
  for (let i = index - 1; i >= 0; i -= 1) {
    if (mask.charAt(i) === MASKED) continue;
    const ch = text.charAt(i);
    if (ch === ' ' || ch === '\t') continue;
    return ch;
  }
  return '';
}

/* ------------------------------------------------------------------ */
/* 可选：重排同一作用域的纯 Dim 区域                                    */
/* ------------------------------------------------------------------ */

function reorderDimRuns(lines: SourceLine[], start: number, end: number): void {
  let i = start;
  while (i <= end) {
    const head = lines[i];
    const firstName = head === undefined ? undefined : dimName(head.text);
    if (firstName === undefined) {
      i += 1;
      continue;
    }
    const names: string[] = [firstName];
    let j = i + 1;
    while (j <= end) {
      const line = lines[j];
      const name = line === undefined ? undefined : dimName(line.text);
      if (name === undefined) break;
      names.push(name);
      j += 1;
    }
    if (j - i >= 2 && new Set(names.map((name) => name.toLowerCase())).size === names.length) {
      const order = names.map((name, index) => ({ key: name.toLowerCase(), index }));
      order.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : a.index - b.index));
      const texts = lines.slice(i, j).map((line) => line.text);
      for (let k = 0; k < order.length; k += 1) {
        const target = lines[i + k];
        const source = texts[order[k]?.index ?? k];
        if (target !== undefined && source !== undefined) target.text = source;
      }
    }
    i = j;
  }
}

function dimName(text: string): string | undefined {
  const mask = maskLine(text, false);
  // 行尾注释随行移动，不影响语义；未闭合字符串一律不参与重排。
  if (mask.startsInString || mask.endsInString) return undefined;
  if (endsWithLineContinuation(mask.mask)) return undefined;
  const code = toCode(mask.mask).trim();
  if (!/^dim\s+/i.test(code)) return undefined;
  if (hasTopLevelPunctuation(code)) return undefined;
  if (/\bnew\b/i.test(code)) return undefined;
  const match = /^dim\s+(?:withevents\s+)?([A-Za-z_][A-Za-z0-9_]*)/i.exec(code);
  return match?.[1];
}

function hasTopLevelPunctuation(code: string): boolean {
  let depth = 0;
  for (let i = 0; i < code.length; i += 1) {
    const ch = code.charAt(i);
    if (ch === '(') depth += 1;
    else if (ch === ')') depth = Math.max(0, depth - 1);
    else if (depth === 0 && (ch === ':' || ch === '=')) return true;
  }
  return false;
}

/* ------------------------------------------------------------------ */
/* 可选：安全逗号处显式续行                                            */
/* ------------------------------------------------------------------ */

const WRAP_BLOCKED =
  /^(?:dim|const|declare|attribute|option|public|private|friend|global|static|type|enum|sub|function|property|event|implements|if|then|else|elseif|end|for|each|next|do|loop|while|wend|with|select|case|on|exit|goto|gosub|stop|resume|randomize|redim|erase|open|close|print|input|write|line|lock|unlock|seek|get|put|let|set)\b/i;

function applyLineWrap(
  lines: SourceLine[],
  plans: LinePlan[],
  settings: ResolvedOptions,
  start: number,
  end: number,
): SourceLine[] {
  const fallbackEol = lines.find((line) => line.eol !== '')?.eol ?? '\n';
  const result: SourceLine[] = [];

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const plan = plans[i];
    if (line === undefined) continue;

    const inRange = i >= start && i <= end;
    if (!inRange || plan === undefined || plan.kind !== 'code' || !plan.statementStart) {
      result.push(line);
      continue;
    }
    if (line.text.length <= LINE_WRAP_LENGTH) {
      result.push(line);
      continue;
    }

    const mask = maskLine(line.text, false);
    if (mask.hasComment || mask.startsInString || mask.endsInString) {
      result.push(line);
      continue;
    }
    if (endsWithLineContinuation(mask.mask)) {
      result.push(line);
      continue;
    }
    const body = line.text.replace(/^[ \t]+/, '');
    if (WRAP_BLOCKED.test(body)) {
      result.push(line);
      continue;
    }

    const pieces = wrapLine(line.text, mask.mask, settings.unit);
    if (pieces === undefined) {
      result.push(line);
      continue;
    }
    for (let k = 0; k < pieces.length; k += 1) {
      const isLast = k === pieces.length - 1;
      result.push({ text: pieces[k] ?? '', eol: isLast ? line.eol : line.eol === '' ? fallbackEol : line.eol });
    }
  }
  return result;
}

function wrapLine(text: string, mask: string, unit: string): string[] | undefined {
  const indent = /^[ \t]*/.exec(text)?.[0] ?? '';
  let depth = 0;
  let topLevelEquals = false;
  let topLevelColon = false;
  const nestedCommas: number[] = [];
  const topCommas: number[] = [];

  for (let i = 0; i < text.length; i += 1) {
    if (mask.charAt(i) === MASKED) continue;
    const ch = text.charAt(i);
    if (ch === '(') depth += 1;
    else if (ch === ')') depth = Math.max(0, depth - 1);
    else if (ch === ',') {
      if (depth > 0) nestedCommas.push(i + 1);
      else topCommas.push(i + 1);
    } else if (ch === '=' && depth === 0) topLevelEquals = true;
    else if (ch === ':' && depth === 0 && text.charAt(i + 1) !== '=') topLevelColon = true;
  }

  const allowTopLevel = !topLevelEquals && !topLevelColon;
  const breaks = [...nestedCommas, ...(allowTopLevel ? topCommas : [])].sort((a, b) => a - b);
  if (breaks.length === 0) return undefined;

  const pieces: string[] = [];
  let cursor = 0;
  let currentIndent = indent;
  let guard = 0;

  while (guard < 64) {
    guard += 1;
    const rest = text.slice(cursor).replace(/^[ \t]+/, '');
    if (currentIndent.length + rest.length <= LINE_WRAP_LENGTH) {
      pieces.push(currentIndent + rest);
      break;
    }
    const limit = LINE_WRAP_LENGTH - currentIndent.length - 2;
    if (limit <= 0) return undefined;
    let chosen = -1;
    for (const point of breaks) {
      if (point <= cursor) continue;
      if (point - cursor > limit) break;
      chosen = point;
    }
    if (chosen < 0) return undefined;
    const segment = text.slice(cursor, chosen).replace(/^[ \t]+/, '').replace(/[ \t]+$/, '');
    pieces.push(`${currentIndent}${segment} _`);
    cursor = chosen;
    while (cursor < text.length) {
      const ch = text.charAt(cursor);
      if (ch !== ' ' && ch !== '\t') break;
      cursor += 1;
    }
    currentIndent = indent + unit;
  }

  return pieces.length > 1 ? pieces : undefined;
}
