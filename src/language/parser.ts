import type { Diagnostic, FoldingRange, Range } from 'vscode-languageserver/node';
import type {
  ModuleKind,
  ParseOptions,
  ParsedModule,
  VbaParameter,
  VbaSymbol,
  VbaToken,
} from './model.js';

/**
 * VBA 6/7 词法关键字（全部小写）。只包含语言关键字与内建类型名，
 * 不包含工程或类型库里的成员名。
 */
export const KEYWORDS: Set<string> = new Set([
  'addressof', 'alias', 'and', 'any', 'as', 'attribute',
  'boolean', 'byref', 'byte', 'byval',
  'call', 'case', 'cdecl', 'close', 'compare', 'const', 'currency',
  'date', 'declare', 'decimal', 'defbool', 'defbyte', 'defcur', 'defdate',
  'defdbl', 'defint', 'deflng', 'deflnglng', 'deflngptr', 'defobj', 'defsng',
  'defstr', 'defvar', 'dim', 'do', 'double',
  'each', 'else', 'elseif', 'empty', 'end', 'enum', 'eqv', 'erase', 'error',
  'event', 'exit', 'explicit',
  'false', 'for', 'friend', 'function',
  'get', 'global', 'gosub', 'goto',
  'if', 'implements', 'imp', 'in', 'input', 'integer', 'is',
  'let', 'lib', 'like', 'line', 'lock', 'long', 'longlong', 'longptr',
  'loop', 'lset',
  'me', 'mod', 'module',
  'new', 'next', 'nothing', 'not', 'null',
  'object', 'on', 'open', 'option', 'optional', 'or',
  'paramarray', 'preserve', 'print', 'private', 'property', 'ptrsafe',
  'public', 'put',
  'raiseevent', 'randomize', 'read', 'redim', 'rem', 'resume', 'return',
  'rset',
  'seek', 'select', 'set', 'single', 'static', 'step', 'stop', 'string',
  'sub',
  'then', 'to', 'true', 'type', 'typeof',
  'unload', 'until',
  'variant', 'wend', 'while', 'with', 'withevents', 'write',
  'xor',
]);

type Access = 'public' | 'private' | 'friend';

const UNKNOWN = Symbol('unknown');
type CondScalar = boolean | number | string;
type CondValue = CondScalar | typeof UNKNOWN;

const TYPE_SUFFIX: Record<string, string> = {
  '%': 'Integer',
  '&': 'Long',
  '!': 'Single',
  '#': 'Double',
  '@': 'Currency',
  '$': 'String',
  '^': 'LongLong',
};

const DATE_LITERAL =
  /^#\s*(?:(?:\d{1,4}[\/\-.]\d{1,2}(?:[\/\-.]\d{1,4})?(?:\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)?)|(?:\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)|(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2},?\s+\d{2,4}))\s*#/i;

const NUMBER_LITERAL =
  /^(?:&[Hh][0-9A-Fa-f]+|&[Oo][0-7]+|&[Bb][01]+|(?:\d+(?:\.\d*)?|\.\d+)(?:[eEdD][+-]?\d+)?)(?:[%&!#@^])?/;

const IDENT_START = /[\p{L}]/u;
const IDENT_PART = /[\p{L}\p{N}_]/u;

/** 合法的单字符运算符。不在其中的独立字符报词法错误。 */
const OPERATOR_CHARS = new Set(['+', '-', '*', '/', '\\', '&', '^', '=', '<', '>', '(', ')', ':', ',', ';', '.', '!', '_', '#']);

/**
 * 结构性保留字：不可用作声明名称（LEX-13）。
 * 依据 MS-VBAL 第 3.3 节：`Line`、`Name`、`Time`、`Seek`、`Print` 等 I/O 与函数词
 * 是上下文关键字，允许作标识符，因此不在此表。
 */
const RESERVED_STRONG: Set<string> = new Set([
  'addressof', 'alias', 'and', 'as', 'boolean', 'byref', 'byval', 'call', 'case', 'const',
  'currency', 'date', 'declare', 'decimal', 'defbool', 'defbyte', 'defcur', 'defdate', 'defdbl',
  'defint', 'deflng', 'deflnglng', 'deflngptr', 'defobj', 'defsng', 'defstr', 'defvar',
  'dim', 'do', 'double', 'each', 'else', 'elseif', 'end', 'enum', 'eqv', 'exit', 'explicit',
  'false', 'for', 'friend', 'function', 'get', 'global', 'gosub', 'goto', 'if', 'implements',
  'imp', 'in', 'integer', 'is', 'let', 'lib', 'long', 'longlong', 'longptr', 'loop', 'me',
  'mod', 'new', 'next', 'not', 'nothing', 'null', 'object', 'on', 'option', 'optional', 'or',
  'paramarray', 'preserve', 'private', 'property', 'public', 'ptrsafe', 'raiseevent', 'redim',
  'rem', 'resume', 'return', 'select', 'set', 'single', 'static', 'step', 'stop', 'string',
  'sub', 'then', 'to', 'true', 'type', 'typeof', 'until', 'variant', 'wend', 'while', 'with',
  'withevents', 'xor',
]);

const SEVERITY_ERROR = 1;
const SEVERITY_WARNING = 2;
const SEVERITY_INFORMATION = 3;

function lower(text: string | undefined): string {
  return text === undefined ? '' : text.toLowerCase();
}

function range(line: number, start: number, endLine: number, end: number): Range {
  return { start: { line, character: start }, end: { line: endLine, character: end } };
}

function tokenRange(token: VbaToken): Range {
  return { start: { line: token.line, character: token.start }, end: { line: token.line, character: token.end } };
}

function makeDiagnostic(
  startLine: number,
  startChar: number,
  endLine: number,
  endChar: number,
  severity: Diagnostic['severity'],
  code: string,
  message: string,
): Diagnostic {
  return {
    range: range(startLine, startChar, endLine, endChar),
    severity,
    code,
    source: 'vba',
    message,
  };
}

function cpAt(text: string, index: number): { ch: string; width: number } {
  const code = text.codePointAt(index);
  if (code === undefined) {
    return { ch: '', width: 0 };
  }
  const ch = String.fromCodePoint(code);
  return { ch, width: ch.length };
}

function isIdentifierStart(ch: string): boolean {
  return ch.length > 0 && IDENT_START.test(ch);
}

function isIdentifierPart(ch: string): boolean {
  return ch.length > 0 && IDENT_PART.test(ch);
}

function readString(line: string, start: number): { end: number; closed: boolean } {
  let i = start + 1;
  while (i < line.length) {
    if (line[i] === '"') {
      if (line[i + 1] === '"') {
        i += 2;
        continue;
      }
      return { end: i + 1, closed: true };
    }
    i += 1;
  }
  return { end: line.length, closed: false };
}

export interface LexedLine {
  tokens: VbaToken[];
  diagnostics: Diagnostic[];
}

/** 词法分析一行物理文本。列号使用 UTF-16 代码单元。 */
export function lexLine(line: string, lineNumber: number): LexedLine {
  const tokens: VbaToken[] = [];
  const diagnostics: Diagnostic[] = [];
  let i = 0;
  let statementStart = true;

  // MS 官方限制：单个物理行最多 1023 字符。
  if (line.length > 1023) {
    diagnostics.push(
      makeDiagnostic(lineNumber, 1023, lineNumber, line.length, SEVERITY_ERROR, 'lex.lineTooLong', '物理行超过 1023 字符。'),
    );
  }

  while (i < line.length) {
    const raw = line[i] ?? '';
    if (raw === ' ' || raw === '\t' || raw === '\u00a0' || raw === '\ufeff') {
      i += 1;
      continue;
    }
    const start = i;

    if (raw === "'") {
      tokens.push({ text: line.slice(i), kind: 'comment', line: lineNumber, start, end: line.length });
      break;
    }

    if (statementStart && (raw === 'R' || raw === 'r')) {
      const rest = line.slice(i);
      if (/^rem(?![A-Za-z0-9_])/i.test(rest)) {
        tokens.push({ text: rest, kind: 'comment', line: lineNumber, start, end: line.length });
        break;
      }
    }

    if (raw === '"') {
      const { end, closed } = readString(line, i);
      tokens.push({ text: line.slice(i, end), kind: 'string', line: lineNumber, start, end });
      if (!closed) {
        diagnostics.push(
          makeDiagnostic(
            lineNumber,
            start,
            lineNumber,
            line.length,
            SEVERITY_ERROR,
            'lex.unterminatedString',
            '字符串缺少结束引号。',
          ),
        );
      }
      i = end;
      statementStart = false;
      continue;
    }

    if (raw === '#') {
      const match = DATE_LITERAL.exec(line.slice(i));
      if (match) {
        const text = match[0];
        tokens.push({ text, kind: 'number', line: lineNumber, start, end: start + text.length });
        i += text.length;
      } else {
        tokens.push({ text: '#', kind: 'operator', line: lineNumber, start, end: start + 1 });
        i += 1;
      }
      statementStart = false;
      continue;
    }

    if (raw === '[') {
      const close = line.indexOf(']', i + 1);
      if (close > i + 1) {
        tokens.push({
          text: line.slice(i + 1, close),
          kind: 'identifier',
          line: lineNumber,
          start,
          end: close + 1,
          escaped: true,
        });
        i = close + 1;
        statementStart = false;
        continue;
      }
    }

    const numeric = NUMBER_LITERAL.exec(line.slice(i));
    if (numeric && raw !== '&' && (/[0-9]/.test(raw) || raw === '.')) {
      const text = numeric[0];
      tokens.push({ text, kind: 'number', line: lineNumber, start, end: start + text.length });
      i += text.length;
      statementStart = false;
      continue;
    }
    if (numeric && raw === '&' && numeric[0].length > 2) {
      const text = numeric[0];
      tokens.push({ text, kind: 'number', line: lineNumber, start, end: start + text.length });
      i += text.length;
      statementStart = false;
      continue;
    }
    if (raw === '.' && /[0-9]/.test(line[i + 1] ?? '')) {
      const text = NUMBER_LITERAL.exec(line.slice(i))?.[0] ?? '.';
      tokens.push({ text, kind: 'number', line: lineNumber, start, end: start + text.length });
      i += text.length;
      statementStart = false;
      continue;
    }

    const { ch: cp } = cpAt(line, i);
    if (isIdentifierStart(cp) || (raw === '_' && isIdentifierPart(line[i + 1] ?? ''))) {
      let j = i + 1;
      while (j < line.length) {
        const { ch } = cpAt(line, j);
        if (!isIdentifierPart(ch)) {
          break;
        }
        j += ch.length;
      }
      const next = line[j] ?? '';
      if ((next === '$' || next === '%' || next === '&' || next === '@' || next === '^') &&
          !isIdentifierPart(line[j + 1] ?? '')) {
        j += 1;
      }
      const text = line.slice(i, j);
      tokens.push({
        text,
        kind: KEYWORDS.has(lower(text)) ? 'keyword' : 'identifier',
        line: lineNumber,
        start,
        end: j,
      });
      i = j;
      statementStart = false;
      continue;
    }

    const two = line.slice(i, i + 2);
    if (two === ':=' || two === '<=' || two === '>=' || two === '<>') {
      tokens.push({ text: two, kind: 'operator', line: lineNumber, start, end: start + 2 });
      i += 2;
      statementStart = false;
      continue;
    }

    // 非法字符（如 @@@ 单独出现、?）：收集连续段报一条，避免雪崩。
    if (!OPERATOR_CHARS.has(raw)) {
      let j = i;
      while (j < line.length) {
        const ch = line[j] ?? '';
        if (ch === ' ' || ch === '\t' || OPERATOR_CHARS.has(ch) || isIdentifierStart(ch) || /[0-9]/.test(ch) || ch === '"' || ch === '\'') { break; }
        j += 1;
      }
      diagnostics.push(
        makeDiagnostic(lineNumber, start, lineNumber, Math.max(j, start + 1), SEVERITY_ERROR, 'lex.invalidCharacter', `存在 VBA 词法中不合法的字符：“${line.slice(i, j)}”。`),
      );
      i = Math.max(j, i + 1);
      continue;
    }

    tokens.push({ text: raw, kind: 'operator', line: lineNumber, start, end: start + 1 });
    i += 1;
    statementStart = raw === ':';
  }

  // 行连接符后不得跟注释（LEX-03）。
  const last = tokens[tokens.length - 1];
  const beforeLast = tokens[tokens.length - 2];
  if (last?.kind === 'comment' && beforeLast?.kind === 'operator' && beforeLast.text === '_') {
    diagnostics.push(
      makeDiagnostic(lineNumber, beforeLast.start, lineNumber, line.length, SEVERITY_ERROR, 'lex.continuationComment', '行连接符后不能跟注释。'),
    );
  }

  return { tokens, diagnostics };
}

/** 对整个文档做词法分析，返回按行顺序排列的 token。 */
export function tokenize(text: string): VbaToken[] {
  const lines = text.split(/\r\n|\n|\r/);
  const result: VbaToken[] = [];
  for (let i = 0; i < lines.length; i += 1) {
    result.push(...lexLine(lines[i] ?? '', i).tokens);
  }
  return result;
}

interface LogicalLine {
  tokens: VbaToken[];
  startLine: number;
  endLine: number;
}

function isContinuation(tokens: VbaToken[]): boolean {
  const last = tokens[tokens.length - 1];
  return last !== undefined && last.kind === 'operator' && last.text === '_';
}

function buildLogicalLines(lines: string[]): LogicalLine[] {
  const groups: LogicalLine[] = [];
  let pending: VbaToken[] | undefined;
  let startLine = 0;
  let lastLine = 0;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i] ?? '';
    const tokens = lexLine(line, i).tokens.filter((token) => token.kind !== 'comment');
    if (tokens.length === 0) {
      if (pending !== undefined) {
        groups.push({ tokens: pending, startLine, endLine: lastLine });
        pending = undefined;
      }
      continue;
    }

    const continued = isContinuation(tokens);
    const usable = continued ? tokens.slice(0, -1) : tokens;
    if (pending === undefined) {
      pending = [];
      startLine = i;
    }
    pending.push(...usable);
    lastLine = i;

    if (!continued) {
      groups.push({ tokens: pending, startLine, endLine: i });
      pending = undefined;
    }
  }
  if (pending !== undefined) {
    groups.push({ tokens: pending, startLine, endLine: lastLine });
  }
  return groups;
}

class CondExpressionParser {
  private position = 0;

  constructor(
    private readonly tokens: VbaToken[],
    private readonly env: Map<string, CondValue>,
  ) {}

  parse(): CondValue {
    const value = this.parseImp();
    return this.position >= this.tokens.length ? value : UNKNOWN;
  }

  private peek(): string {
    return lower(this.tokens[this.position]?.text);
  }

  private consume(): VbaToken | undefined {
    const token = this.tokens[this.position];
    this.position += 1;
    return token;
  }

  private parseImp(): CondValue {
    let left = this.parseEqv();
    while (this.peek() === 'imp') {
      this.consume();
      left = this.binary(left, this.parseEqv(), 'imp');
    }
    return left;
  }

  private parseEqv(): CondValue {
    let left = this.parseXor();
    while (this.peek() === 'eqv') {
      this.consume();
      left = this.binary(left, this.parseXor(), 'eqv');
    }
    return left;
  }

  private parseXor(): CondValue {
    let left = this.parseOr();
    while (this.peek() === 'xor') {
      this.consume();
      left = this.binary(left, this.parseOr(), 'xor');
    }
    return left;
  }

  private parseOr(): CondValue {
    let left = this.parseAnd();
    while (this.peek() === 'or') {
      this.consume();
      left = this.binary(left, this.parseAnd(), 'or');
    }
    return left;
  }

  private parseAnd(): CondValue {
    let left = this.parseNot();
    while (this.peek() === 'and') {
      this.consume();
      left = this.binary(left, this.parseNot(), 'and');
    }
    return left;
  }

  private parseNot(): CondValue {
    if (this.peek() === 'not') {
      this.consume();
      return this.not(this.parseNot());
    }
    return this.parseComparison();
  }

  private parseComparison(): CondValue {
    const left = this.parseAdditive();
    const op = this.peek();
    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      this.consume();
      return this.compare(left, this.parseAdditive(), op);
    }
    return left;
  }

  private parseAdditive(): CondValue {
    let left = this.parseMultiplicative();
    for (;;) {
      const op = this.peek();
      if (op !== '+' && op !== '-' && op !== '&') {
        return left;
      }
      this.consume();
      left = this.binary(left, this.parseMultiplicative(), op);
    }
  }

  private parseMultiplicative(): CondValue {
    let left = this.parseUnary();
    for (;;) {
      const op = this.peek();
      if (op !== '*' && op !== '/' && op !== '\\' && op !== 'mod') {
        return left;
      }
      this.consume();
      left = this.binary(left, this.parseUnary(), op);
    }
  }

  private parseUnary(): CondValue {
    const op = this.peek();
    if (op === '-' || op === '+') {
      this.consume();
      const value = this.parseUnary();
      if (value === UNKNOWN || typeof value !== 'number') {
        return UNKNOWN;
      }
      return op === '-' ? -value : value;
    }
    return this.parsePrimary();
  }

  private parsePrimary(): CondValue {
    const token = this.tokens[this.position];
    if (token === undefined) {
      return UNKNOWN;
    }
    if (token.text === '(') {
      this.consume();
      const value = this.parseImp();
      if (this.peek() === ')') {
        this.consume();
        return value;
      }
      return UNKNOWN;
    }
    if (token.kind === 'string') {
      this.consume();
      return unquoteVbaString(token.text);
    }
    if (token.kind === 'number') {
      this.consume();
      const value = parseNumericText(token.text);
      return value === undefined ? UNKNOWN : value;
    }
    const word = lower(token.text);
    if (word === 'true') {
      this.consume();
      return true;
    }
    if (word === 'false') {
      this.consume();
      return false;
    }
    if (token.kind === 'identifier' || token.kind === 'keyword') {
      this.consume();
      if (word === 'defined' && this.peek() === '(') {
        this.consume();
        const nameToken = this.consume();
        const name = lower(nameToken?.text);
        if (this.peek() === ')') {
          this.consume();
        }
        return name !== '' && this.env.has(name);
      }
      const value = this.env.get(word);
      return value === undefined ? UNKNOWN : value;
    }
    return UNKNOWN;
  }

  private binary(left: CondValue, right: CondValue, op: string): CondValue {
    if (left === UNKNOWN || right === UNKNOWN) {
      return UNKNOWN;
    }
    if (op === '&') {
      return `${String(left)}${String(right)}`;
    }
    if (op === 'eqv' || op === 'imp') {
      const a = truthy(left);
      const b = truthy(right);
      if (a === undefined || b === undefined) {
        return UNKNOWN;
      }
      return op === 'eqv' ? a === b : !a || b;
    }
    if (typeof left === 'boolean' || typeof right === 'boolean') {
      const a = truthy(left);
      const b = truthy(right);
      if (a === undefined || b === undefined) {
        return UNKNOWN;
      }
      if (op === 'and') {
        return a && b;
      }
      if (op === 'or') {
        return a || b;
      }
      if (op === 'xor') {
        return a !== b;
      }
      return UNKNOWN;
    }
    if (typeof left === 'number' && typeof right === 'number') {
      switch (op) {
        case '+': return left + right;
        case '-': return left - right;
        case '*': return left * right;
        case '/': return right === 0 ? UNKNOWN : left / right;
        case '\\': return right === 0 ? UNKNOWN : Math.trunc(left / right);
        case 'mod': return right === 0 ? UNKNOWN : left % right;
        case 'and': return left & right;
        case 'or': return left | right;
        case 'xor': return left ^ right;
        default: return UNKNOWN;
      }
    }
    if (typeof left === 'string' && typeof right === 'string' && op === '+') {
      return left + right;
    }
    return UNKNOWN;
  }

  private compare(left: CondValue, right: CondValue, op: string): CondValue {
    if (left === UNKNOWN || right === UNKNOWN) {
      return UNKNOWN;
    }
    let a: number | string;
    let b: number | string;
    if (typeof left === 'number' && typeof right === 'number') {
      a = left;
      b = right;
    } else if (typeof left === 'string' && typeof right === 'string') {
      a = left;
      b = right;
    } else {
      return UNKNOWN;
    }
    switch (op) {
      case '=': return a === b;
      case '<>': return a !== b;
      case '<': return a < b;
      case '>': return a > b;
      case '<=': return a <= b;
      case '>=': return a >= b;
      default: return UNKNOWN;
    }
  }

  private not(value: CondValue): CondValue {
    if (value === UNKNOWN) {
      return UNKNOWN;
    }
    if (typeof value === 'boolean') {
      return !value;
    }
    if (typeof value === 'number') {
      return ~value;
    }
    return UNKNOWN;
  }
}

function truthy(value: CondScalar): boolean | undefined {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  return undefined;
}

function unquoteVbaString(text: string): string {
  const body = text.startsWith('"') ? text.slice(1) : text;
  const withoutEnd = body.endsWith('"') ? body.slice(0, -1) : body;
  return withoutEnd.replace(/""/g, '"');
}

function parseNumericText(text: string): number | undefined {
  let value = text.trim().replace(/[%&!#@^]$/, '');
  if (/^&[Hh]/.test(value)) {
    return Number.parseInt(value.slice(2), 16);
  }
  if (/^&[Oo]/.test(value)) {
    return Number.parseInt(value.slice(2), 8);
  }
  if (/^&[Bb]/.test(value)) {
    return Number.parseInt(value.slice(2), 2);
  }
  value = value.replace(/[dD]/, 'e');
  if (!/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value)) {
    return undefined;
  }
  return Number(value);
}

interface CondFrame {
  parentActive: boolean;
  active: boolean;
  knownTaken: boolean;
  unknownSeen: boolean;
  startLine: number;
}

function directiveWord(tokens: VbaToken[]): string {
  return lower(tokens[1]?.text);
}

function isDirective(group: LogicalLine): boolean {
  const first = group.tokens[0];
  if (first === undefined || first.text !== '#') {
    return false;
  }
  const word = directiveWord(group.tokens);
  return word === 'if' || word === 'elseif' || word === 'else' || word === 'end' || word === 'const';
}

interface ConditionalResult {
  lineActive: boolean[];
  unknownConditional: boolean;
  diagnostics: Diagnostic[];
}

function evaluateConditionals(
  groups: LogicalLine[],
  lineCount: number,
  protectedLines: Set<number>,
  options: ParseOptions,
): ConditionalResult {
  const lineActive = new Array<boolean>(lineCount).fill(true);
  const diagnostics: Diagnostic[] = [];
  const frames: CondFrame[] = [];
  let unknownConditional = false;

  const env = new Map<string, CondValue>();
  const bits = options.targetBits ?? 64;
  env.set('vba7', true);
  env.set('vba6', false);
  env.set('win32', bits === 32);
  env.set('win64', bits === 64);
  env.set('win16', false);
  env.set('mac', false);
  for (const [name, value] of Object.entries(options.constants ?? {})) {
    env.set(name.toLowerCase(), value);
  }

  const activeNow = (): boolean => (frames.length === 0 ? true : frames[frames.length - 1]!.active);
  const parentActiveNow = (): boolean => (frames.length === 0 ? true : frames[frames.length - 1]!.active);

  const markUnknown = (group: LogicalLine): void => {
    unknownConditional = true;
    const first = group.tokens[0]!;
    const last = group.tokens[group.tokens.length - 1] ?? first;
    diagnostics.push(
      makeDiagnostic(
        group.startLine,
        first.start,
        last.line,
        last.end,
        SEVERITY_INFORMATION,
        'cond.unknown',
        '条件编译常量未定义，无法判定该分支；该分支已从语义分析中排除。',
      ),
    );
  };

  for (const group of groups) {
    if (isDirective(group)) {
      const word = directiveWord(group.tokens);
      if (word === 'if') {
        const parentActive = activeNow();
        const expressionTokens = group.tokens.slice(2);
        let thenIndex = expressionTokens.findIndex((token) => lower(token.text) === 'then');
        if (thenIndex < 0) {
          thenIndex = expressionTokens.length;
        }
        const expression = expressionTokens.slice(0, thenIndex);
        if (expression.length === 0) {
          diagnostics.push(
            makeDiagnostic(
              group.startLine,
              group.tokens[0]!.start,
              group.endLine,
              (group.tokens[group.tokens.length - 1] ?? group.tokens[0]!).end,
              SEVERITY_ERROR,
              'cond.missingExpression',
              '#If 缺少条件表达式。',
            ),
          );
        }
        if (!parentActive) {
          frames.push({ parentActive: false, active: false, knownTaken: true, unknownSeen: false, startLine: group.startLine });
          continue;
        }
        const value = expression.length === 0 ? UNKNOWN : new CondExpressionParser(expression, env).parse();
        if (value === UNKNOWN) {
          frames.push({ parentActive: true, active: false, knownTaken: true, unknownSeen: true, startLine: group.startLine });
          markUnknown(group);
        } else {
          const truth = truthy(value);
          if (truth === undefined) {
            frames.push({ parentActive: true, active: false, knownTaken: true, unknownSeen: true, startLine: group.startLine });
            markUnknown(group);
          } else {
            frames.push({ parentActive: true, active: truth, knownTaken: truth, unknownSeen: false, startLine: group.startLine });
          }
        }
        continue;
      }

      if (word === 'elseif') {
        const frame = frames[frames.length - 1];
        if (frame === undefined) {
          diagnostics.push(
            makeDiagnostic(
              group.startLine,
              group.tokens[0]!.start,
              group.endLine,
              (group.tokens[group.tokens.length - 1] ?? group.tokens[0]!).end,
              SEVERITY_ERROR,
              'cond.unmatchedElseIf',
              '#ElseIf 没有对应的 #If。',
            ),
          );
          continue;
        }
        let expressionTokens = group.tokens.slice(2);
        let thenIndex = expressionTokens.findIndex((token) => lower(token.text) === 'then');
        if (thenIndex < 0) {
          thenIndex = expressionTokens.length;
        }
        expressionTokens = expressionTokens.slice(0, thenIndex);
        if (!frame.parentActive || frame.knownTaken) {
          frame.active = false;
          continue;
        }
        if (frame.unknownSeen) {
          frame.active = false;
          markUnknown(group);
          continue;
        }
        const value = expressionTokens.length === 0 ? UNKNOWN : new CondExpressionParser(expressionTokens, env).parse();
        const truth = value === UNKNOWN ? undefined : truthy(value);
        if (truth === undefined) {
          frame.unknownSeen = true;
          frame.active = false;
          markUnknown(group);
        } else {
          frame.active = truth;
          frame.knownTaken = truth;
        }
        continue;
      }

      if (word === 'else') {
        const frame = frames[frames.length - 1];
        if (frame === undefined) {
          diagnostics.push(
            makeDiagnostic(
              group.startLine,
              group.tokens[0]!.start,
              group.endLine,
              (group.tokens[group.tokens.length - 1] ?? group.tokens[0]!).end,
              SEVERITY_ERROR,
              'cond.unmatchedElse',
              '#Else 没有对应的 #If。',
            ),
          );
          continue;
        }
        if (!frame.parentActive || frame.knownTaken) {
          frame.active = false;
          continue;
        }
        if (frame.unknownSeen) {
          frame.active = false;
          markUnknown(group);
          continue;
        }
        frame.active = true;
        frame.knownTaken = true;
        continue;
      }

      if (word === 'const') {
        if (!activeNow()) {
          continue;
        }
        const nameToken = group.tokens[2];
        const equals = group.tokens.findIndex((token, index) => index > 2 && token.text === '=');
        if (nameToken === undefined || nameToken.kind !== 'identifier' || equals < 0) {
          diagnostics.push(
            makeDiagnostic(
              group.startLine,
              group.tokens[0]!.start,
              group.endLine,
              (group.tokens[group.tokens.length - 1] ?? group.tokens[0]!).end,
              SEVERITY_ERROR,
              'cond.malformedConst',
              '#Const 需要名称与赋值表达式。',
            ),
          );
          continue;
        }
        const expression = group.tokens.slice(equals + 1);
        const value = expression.length === 0 ? UNKNOWN : new CondExpressionParser(expression, env).parse();
        env.set(lower(nameToken.text), value);
        continue;
      }

      // #End If
      if (word === 'end') {
        if (lower(group.tokens[2]?.text) !== 'if') {
          continue;
        }
        if (frames.length === 0) {
          diagnostics.push(
            makeDiagnostic(
              group.startLine,
              group.tokens[0]!.start,
              group.endLine,
              (group.tokens[group.tokens.length - 1] ?? group.tokens[0]!).end,
              SEVERITY_ERROR,
              'cond.unmatchedEndIf',
              '#End If 没有对应的 #If。',
            ),
          );
        } else {
          frames.pop();
        }
        continue;
      }
      continue;
    }

    if (!activeNow()) {
      for (let line = group.startLine; line <= group.endLine; line += 1) {
        if (!protectedLines.has(line)) {
          lineActive[line] = false;
        }
      }
    }
  }

  for (const frame of frames) {
    diagnostics.push(
      makeDiagnostic(
        frame.startLine,
        0,
        frame.startLine,
        1,
        SEVERITY_ERROR,
        'cond.unclosedIf',
        '条件编译块未闭合，缺少 #End If。',
      ),
    );
  }

  return { lineActive, unknownConditional, diagnostics };
}

interface HeaderInfo {
  protectedLines: Set<number>;
  attributes: Map<string, string>;
  descriptions: Map<string, string>;
  nameLocation?: { line: number; start: number; end: number };
  predeclaredId?: boolean;
  versionClass: boolean;
  versionForm: boolean;
  designStart?: number;
  designEnd?: number;
}

function unquoteAttributeValue(raw: string): string {
  const text = raw.trim();
  if (!text.startsWith('"')) {
    return text;
  }
  let result = '';
  let i = 1;
  while (i < text.length) {
    if (text[i] === '"') {
      if (text[i + 1] === '"') {
        result += '"';
        i += 2;
        continue;
      }
      break;
    }
    result += text[i];
    i += 1;
  }
  return result;
}

const ATTRIBUTE_RE = /^Attribute\s+([A-Za-z_]\w*)(?:\.([A-Za-z_]\w*))?\s*=\s*(.*)$/i;

function scanHeader(lines: string[]): HeaderInfo {
  const info: HeaderInfo = {
    protectedLines: new Set<number>(),
    attributes: new Map<string, string>(),
    descriptions: new Map<string, string>(),
    versionClass: false,
    versionForm: false,
  };

  const recordAttribute = (line: string, lineIndex: number): void => {
    const match = ATTRIBUTE_RE.exec(line);
    if (!match) {
      return;
    }
    const owner = match[1] ?? '';
    const member = match[2];
    const rawValue = match[3] ?? '';
    if (member !== undefined) {
      if (member.toLowerCase() === 'vb_description') {
        info.descriptions.set(owner.toLowerCase(), unquoteAttributeValue(rawValue));
      }
      return;
    }
    const name = owner.toLowerCase();
    const value = unquoteAttributeValue(rawValue);
    info.attributes.set(name, value);
    if (name === 'vb_predeclaredid') {
      info.predeclaredId = /^true$/i.test(value.trim()) || value.trim() === '-1';
    }
    if (name === 'vb_name' && info.nameLocation === undefined) {
      const quoteStart = line.indexOf('"');
      if (quoteStart >= 0) {
        const quoteEnd = line.indexOf('"', quoteStart + 1);
        if (quoteEnd > quoteStart) {
          info.nameLocation = { line: lineIndex, start: quoteStart + 1, end: quoteEnd };
        }
      }
    }
  };

  let i = 0;
  let designDepth = 0;
  while (i < lines.length) {
    const trimmed = (lines[i] ?? '').trim();
    if (trimmed === '') {
      i += 1;
      continue;
    }
    if (/^version\s+/i.test(trimmed)) {
      info.protectedLines.add(i);
      if (/^version\s+5\./i.test(trimmed)) {
        info.versionForm = true;
      }
      if (/^version\s+1\.0\s+class/i.test(trimmed)) {
        info.versionClass = true;
      }
      if (info.designStart === undefined) {
        info.designStart = i;
      }
      info.designEnd = i;
      i += 1;
      continue;
    }
    if (/^begin\b/i.test(trimmed)) {
      const start = i;
      info.protectedLines.add(i);
      info.designStart = info.designStart ?? i;
      designDepth = 1;
      i += 1;
      while (i < lines.length && designDepth > 0) {
        const inner = (lines[i] ?? '').trim();
        info.protectedLines.add(i);
        info.designEnd = i;
        i += 1;
        if (/^begin\b/i.test(inner)) {
          designDepth += 1;
        } else if (/^end$/i.test(inner)) {
          designDepth -= 1;
        }
      }
      if (info.designStart === undefined) {
        info.designStart = start;
      }
      continue;
    }
    if (/^attribute\b/i.test(trimmed)) {
      info.protectedLines.add(i);
      recordAttribute(trimmed, i);
      info.designStart = info.designStart ?? i;
      info.designEnd = i;
      i += 1;
      continue;
    }
    if (trimmed.startsWith("'")) {
      i += 1;
      continue;
    }
    if (/^end$/i.test(trimmed)) {
      info.protectedLines.add(i);
      info.designEnd = i;
      i += 1;
      continue;
    }
    break;
  }

  for (let line = 0; line < lines.length; line += 1) {
    const trimmed = (lines[line] ?? '').trim();
    if (/^attribute\b/i.test(trimmed)) {
      info.protectedLines.add(line);
      recordAttribute(trimmed, line);
    }
  }

  return info;
}

function splitTopLevel(tokens: VbaToken[], separator: string): VbaToken[][] {
  const parts: VbaToken[][] = [];
  let current: VbaToken[] = [];
  let depth = 0;
  for (const token of tokens) {
    if (token.text === '(' || token.text === '[') {
      depth += 1;
    } else if (token.text === ')' || token.text === ']') {
      depth = Math.max(0, depth - 1);
    }
    if (depth === 0 && token.text === separator) {
      parts.push(current);
      current = [];
      continue;
    }
    current.push(token);
  }
  parts.push(current);
  return parts;
}

interface ParsedParameter extends VbaParameter {
  hadType: boolean;
  array: boolean;
  token: VbaToken;
}

function parseParameters(tokens: VbaToken[], openIndex: number): { parameters: ParsedParameter[]; closeIndex: number } {
  const open = tokens[openIndex];
  if (open === undefined || open.text !== '(') {
    return { parameters: [], closeIndex: openIndex - 1 };
  }
  let depth = 0;
  let close = tokens.length;
  for (let i = openIndex; i < tokens.length; i += 1) {
    const text = tokens[i]?.text;
    if (text === '(') {
      depth += 1;
    } else if (text === ')') {
      depth -= 1;
      if (depth === 0) {
        close = i;
        break;
      }
    }
  }
  const inner = tokens.slice(openIndex + 1, close);
  const parameters: ParsedParameter[] = [];
  for (const part of splitTopLevel(inner, ',')) {
    if (part.length === 0) {
      continue;
    }
    let optional = false;
    let byRef = true;
    let paramArray = false;
    let nameToken: VbaToken | undefined;
    for (let i = 0; i < part.length; i += 1) {
      const token = part[i];
      if (token === undefined) {
        continue;
      }
      const word = lower(token.text);
      if (word === 'optional') {
        optional = true;
        continue;
      }
      if (word === 'byval') {
        byRef = false;
        continue;
      }
      if (word === 'byref') {
        byRef = true;
        continue;
      }
      if (word === 'paramarray') {
        paramArray = true;
        continue;
      }
      if (token.kind === 'identifier' || token.kind === 'keyword') {
        if (word === 'as') {
          break;
        }
        nameToken = token;
        break;
      }
    }
    if (nameToken === undefined) {
      continue;
    }
    let baseName = nameToken.text;
    const suffix = baseName.slice(-1);
    if (TYPE_SUFFIX[suffix] !== undefined) {
      baseName = baseName.slice(0, -1);
    }
    let array = false;
    let closeName = -1;
    for (let i = 0; i < part.length; i += 1) {
      const token = part[i];
      if (token === nameToken) {
        closeName = i;
        continue;
      }
      if (closeName >= 0 && token?.text === '(') {
        array = true;
        break;
      }
    }
    let typeName = 'Variant';
    let hadType = false;
    for (let i = 0; i < part.length; i += 1) {
      if (lower(part[i]?.text) === 'as') {
        const typeToken = part[i + 1];
        if (typeToken !== undefined && lower(typeToken.text) !== 'new') {
          typeName = lookupTypeAfterAs(part, i + 1) ?? typeToken.text;
          hadType = true;
        }
        break;
      }
    }
    parameters.push({ name: baseName, typeName, optional, byRef, paramArray, hadType, array, token: nameToken });
  }
  return { parameters, closeIndex: close };
}

function parameterSignature(parameter: ParsedParameter): string {
  const prefix = parameter.paramArray
    ? 'ParamArray '
    : `${parameter.optional ? 'Optional ' : ''}${parameter.byRef ? 'ByRef ' : 'ByVal '}`;
  return `${prefix}${parameter.name}${parameter.array ? '()' : ''} As ${parameter.typeName}`;
}

function publicParameters(parameters: ParsedParameter[]): VbaParameter[] {
  return parameters.map(({ name, typeName, optional, byRef, paramArray, array }) => ({
    name,
    typeName,
    optional,
    byRef,
    ...(paramArray === true ? { paramArray: true } : {}),
    ...(array === true ? { array: true } : {}),
  }));
}

interface DeclPrefix {
  access?: Access;
  global: boolean;
  isStatic: boolean;
  withEvents: boolean;
  isDefault: boolean;
  index: number;
}

function parseDeclPrefix(tokens: VbaToken[]): DeclPrefix {
  let access: Access | undefined;
  let global = false;
  let isStatic = false;
  let withEvents = false;
  let isDefault = false;
  let i = 0;
  for (;;) {
    const word = lower(tokens[i]?.text);
    if (word === 'public') {
      access = 'public';
      i += 1;
      continue;
    }
    if (word === 'private') {
      access = 'private';
      i += 1;
      continue;
    }
    if (word === 'friend') {
      access = 'friend';
      i += 1;
      continue;
    }
    if (word === 'default') {
      // `Public Default Property Get` 等形式；Default 只影响语义，不参与块配对。
      isDefault = true;
      i += 1;
      continue;
    }
    if (word === 'global') {
      access = 'public';
      global = true;
      i += 1;
      continue;
    }
    if (word === 'static') {
      isStatic = true;
      i += 1;
      continue;
    }
    if (word === 'withevents') {
      withEvents = true;
      i += 1;
      continue;
    }
    break;
  }
  return { access, global, isStatic, withEvents, isDefault, index: i };
}

function accessFrom(prefix: DeclPrefix, moduleLevel: boolean): Access {
  if (prefix.access !== undefined) {
    return prefix.access;
  }
  return moduleLevel ? 'public' : 'private';
}

function accessLabel(access: Access): string {
  return access === 'public' ? 'Public' : access === 'private' ? 'Private' : 'Friend';
}

interface BlockEntry {
  type: 'proc' | 'if' | 'with' | 'for' | 'do' | 'while' | 'select' | 'type' | 'enum';
  startLine: number;
  startChar: number;
  symbol?: VbaSymbol;
  procKeyword?: string;
  accessor?: string;
  name?: string;
}

interface ParserContext {
  uri: string;
  moduleName: string;
  moduleId: string;
  lines: string[];
  lineActive: boolean[];
  symbols: VbaSymbol[];
  diagnostics: Diagnostic[];
  folds: FoldingRange[];
  stack: BlockEntry[];
  currentProc?: BlockEntry;
  explicit: boolean;
  optionsSeen: Set<string>;
  declarationSeen: boolean;
  withEventsNames: Set<string>;
  protectedLines: Set<number>;
}

function moduleScope(context: ParserContext): string {
  return context.moduleName;
}

function procScopeKey(context: ParserContext, procName: string, accessor?: string): string {
  return accessor === undefined
    ? `${context.moduleName}.${procName}`
    : `${context.moduleName}.${procName}.${accessor}`;
}

function currentProcedureScopeKey(context: ParserContext): string {
  const proc = context.currentProc;
  if (proc === undefined || proc.name === undefined) {
    return moduleScope(context);
  }
  return procScopeKey(context, proc.name, proc.accessor);
}

function symbolId(context: ParserContext, name: string, scope: string, kind: string): string {
  return `${context.uri}#${name}#${scope}#${kind}`;
}

function collectDocumentation(context: ParserContext, line: number): string | undefined {
  const collected: string[] = [];
  let index = line - 1;
  while (index >= 0) {
    const raw = context.lines[index] ?? '';
    const trimmed = raw.trim();
    if (trimmed === '') {
      break;
    }
    if (context.protectedLines.has(index)) {
      break;
    }
    if (!context.lineActive[index]) {
      break;
    }
    if (trimmed.startsWith("'")) {
      collected.unshift(trimmed.replace(/^'+\s?/, '').trimEnd());
      index -= 1;
      continue;
    }
    if (/^rem(?![A-Za-z0-9_])/i.test(trimmed)) {
      collected.unshift(trimmed.replace(/^rem\b\s?/i, '').trimEnd());
      index -= 1;
      continue;
    }
    break;
  }
  while (collected.length > 0 && collected[0] === '') {
    collected.shift();
  }
  if (collected.length === 0) {
    return undefined;
  }
  return collected.join('\n');
}

function addFold(context: ParserContext, startLine: number, endLine: number, kind: 'region' | 'comment'): void {
  if (endLine <= startLine || startLine < 0) {
    return;
  }
  const capped = Math.min(endLine, context.lines.length - 1);
  if (capped <= startLine) {
    return;
  }
  context.folds.push({ startLine, endLine: capped, kind });
}

function reportMissingEnd(context: ParserContext, entry: BlockEntry): void {
  const label: Record<BlockEntry['type'], string> = {
    proc: entry.procKeyword === 'property' ? 'Property' : entry.procKeyword === 'function' ? 'Function' : 'Sub',
    if: 'If',
    with: 'With',
    for: 'For',
    do: 'Do',
    while: 'While',
    select: 'Select',
    type: 'Type',
    enum: 'Enum',
  };
  context.diagnostics.push(
    makeDiagnostic(
      entry.startLine,
      entry.startChar,
      entry.startLine,
      entry.startChar + 1,
      SEVERITY_ERROR,
      'syntax.missingEnd',
      `缺少 End ${label[entry.type]}，块在此处仍未闭合。`,
    ),
  );
}

function previousLineEnd(context: ParserContext, line: number): { line: number; end: number } {
  const previous = Math.max(0, line - 1);
  return { line: previous, end: (context.lines[previous] ?? '').length };
}

function closeEntry(context: ParserContext, entry: BlockEntry, endLine: number, endChar: number, fold: boolean): void {
  if (entry.symbol !== undefined) {
    entry.symbol.range = range(entry.symbol.range.start.line, entry.symbol.range.start.character, endLine, endChar);
  }
  if (fold) {
    addFold(context, entry.startLine, endLine - 1, 'region');
  }
}

function popBlock(context: ParserContext, endLine: number, endChar: number, fold = true): void {
  const entry = context.stack.pop();
  if (entry === undefined) {
    return;
  }
  closeEntry(context, entry, endLine, endChar, fold);
  if (entry.type === 'proc') {
    context.currentProc = undefined;
  }
}

function closeByType(context: ParserContext, type: BlockEntry['type'], statementLine: number, endChar: number): void {
  let index = -1;
  for (let i = context.stack.length - 1; i >= 0; i -= 1) {
    if (context.stack[i]?.type === type) {
      index = i;
      break;
    }
  }
  if (index < 0) {
    context.diagnostics.push(
      makeDiagnostic(
        statementLine,
        0,
        statementLine,
        Math.max(1, endChar),
        SEVERITY_ERROR,
        'syntax.unexpectedEnd',
        `多余的 End ${type === 'proc' ? '' : type}，没有匹配的开始块。`,
      ),
    );
    return;
  }
  while (context.stack.length - 1 > index) {
    const entry = context.stack[context.stack.length - 1]!;
    const previous = previousLineEnd(context, statementLine);
    reportMissingEnd(context, entry);
    popBlock(context, previous.line, previous.end, false);
  }
  popBlock(context, statementLine, endChar);
}

function closeProcedure(context: ParserContext, expected: string, statementLine: number, endChar: number): void {
  let index = -1;
  for (let i = context.stack.length - 1; i >= 0; i -= 1) {
    if (context.stack[i]?.type === 'proc') {
      index = i;
      break;
    }
  }
  if (index < 0) {
    context.diagnostics.push(
      makeDiagnostic(
        statementLine,
        0,
        statementLine,
        Math.max(1, endChar),
        SEVERITY_ERROR,
        'syntax.unexpectedEnd',
        `多余的 End ${expected}，没有匹配的过程声明。`,
      ),
    );
    return;
  }
  const proc = context.stack[index]!;
  while (context.stack.length - 1 > index) {
    const entry = context.stack[context.stack.length - 1]!;
    const previous = previousLineEnd(context, statementLine);
    reportMissingEnd(context, entry);
    popBlock(context, previous.line, previous.end, false);
  }
  if (proc.procKeyword !== expected) {
    context.diagnostics.push(
      makeDiagnostic(
        statementLine,
        0,
        statementLine,
        Math.max(1, endChar),
        SEVERITY_ERROR,
        'syntax.mismatchedEnd',
        `End ${expected} 与过程声明的 End ${proc.procKeyword ?? ''} 不匹配。`,
      ),
    );
  }
  popBlock(context, statementLine, endChar);
}

function closeUnclosedBeforeProcedure(context: ParserContext, statementLine: number, endChar: number): void {
  let procIndex = -1;
  for (let i = context.stack.length - 1; i >= 0; i -= 1) {
    if (context.stack[i]?.type === 'proc') {
      procIndex = i;
      break;
    }
  }
  if (procIndex >= 0) {
    const previous = previousLineEnd(context, statementLine);
    while (context.stack.length - 1 >= procIndex) {
      const entry = context.stack[context.stack.length - 1]!;
      reportMissingEnd(context, entry);
      popBlock(context, previous.line, previous.end, false);
    }
    return;
  }
  const previous = previousLineEnd(context, statementLine);
  while (context.stack.length > 0 && (context.stack[context.stack.length - 1]?.type === 'type' || context.stack[context.stack.length - 1]?.type === 'enum')) {
    const entry = context.stack[context.stack.length - 1]!;
    reportMissingEnd(context, entry);
    popBlock(context, previous.line, previous.end, false);
  }
}

function addSymbol(context: ParserContext, symbol: VbaSymbol): void {
  context.symbols.push(symbol);
}

function findNameToken(tokens: VbaToken[], start: number): VbaToken | undefined {
  for (let i = start; i < tokens.length; i += 1) {
    const token = tokens[i];
    if (token !== undefined && token.kind === 'identifier') {
      return token;
    }
  }
  return undefined;
}

function stripTypeSuffix(name: string): { name: string; suffixType?: string } {
  const suffix = name.slice(-1);
  if (TYPE_SUFFIX[suffix] !== undefined) {
    return { name: name.slice(0, -1), suffixType: TYPE_SUFFIX[suffix] };
  }
  return { name };
}

function parseVariableDeclaration(
  context: ParserContext,
  statementTokens: VbaToken[],
  startIndex: number,
  options: {
    access: Access;
    moduleLevel: boolean;
    isStatic: boolean;
    withEvents: boolean;
    constant: boolean;
    idScope: string;
    scope?: string;
    container?: string;
  },
): void {
  const tokens = statementTokens.slice(startIndex);
  if (tokens.length === 0) {
    // `Dim` 等声明缺少名称（DIAG-02/06）。
    const anchor = statementTokens[startIndex];
    context.diagnostics.push(
      makeDiagnostic(anchor?.line ?? 0, anchor?.start ?? 0, anchor?.line ?? 0, anchor?.end ?? 1, SEVERITY_ERROR, 'syntax.missingName', '声明缺少名称。'),
    );
    return;
  }
  for (const part of splitTopLevel(tokens, ',')) {
    if (part.length === 0) {
      continue;
    }
    let nameIndex = -1;
    for (let i = 0; i < part.length; i += 1) {
      const token = part[i];
      if (token === undefined) {
        continue;
      }
      if (lower(token.text) === 'withevents') {
        continue;
      }
      if (token.kind === 'identifier' || (token.kind === 'keyword' && !RESERVED_STRONG.has(lower(token.text)))) {
        nameIndex = i;
        break;
      }
    }
    if (nameIndex < 0) {
      const anchor = part[0] ?? statementTokens[startIndex];
      context.diagnostics.push(
        makeDiagnostic(anchor?.line ?? 0, anchor?.start ?? 0, anchor?.line ?? 0, anchor?.end ?? 1, SEVERITY_ERROR, 'syntax.missingName', '声明缺少名称。'),
      );
      continue;
    }
    const nameToken = part[nameIndex]!;
    const { name, suffixType } = stripTypeSuffix(nameToken.text);
    const declaratorStart = part[0] ?? nameToken;
    const declaratorEnd = part[part.length - 1] ?? nameToken;

    let array = false;
    let parenDepth = 0;
    let sawName = false;
    let asIndex = -1;
    for (let i = 0; i < part.length; i += 1) {
      const token = part[i];
      if (token === undefined) {
        continue;
      }
      if (token === nameToken) {
        sawName = true;
        continue;
      }
      if (!sawName) {
        continue;
      }
      if (token.text === '(' && parenDepth === 0) {
        array = true;
        parenDepth += 1;
        continue;
      }
      if (token.text === ')') {
        parenDepth = Math.max(0, parenDepth - 1);
        continue;
      }
      if (parenDepth === 0 && lower(token.text) === 'as' && asIndex < 0) {
        asIndex = i;
      }
    }

    let typeName: string | undefined = suffixType;
    if (asIndex >= 0) {
      for (let i = asIndex + 1; i < part.length; i += 1) {
        const token = part[i];
        if (token === undefined) {
          continue;
        }
        if (lower(token.text) === 'new') {
          continue;
        }
        if (token.kind === 'identifier' || token.kind === 'keyword') {
          typeName = lookupTypeAfterAs(part, i) ?? token.text;
        }
        break;
      }
    }
    if (typeName === undefined && options.constant) {
      let equals = -1;
      for (let i = 0; i < part.length; i += 1) {
        if (part[i]?.text === '=') {
          equals = i;
          break;
        }
      }
      const valueToken = equals >= 0 ? part[equals + 1] : undefined;
      if (valueToken !== undefined) {
        if (valueToken.kind === 'string') {
          typeName = 'String';
        } else if (valueToken.kind === 'number') {
          typeName = valueToken.text.startsWith('#')
            ? 'Date'
            : /[.eEdD]/.test(valueToken.text.replace(/^&[HhOoBb]/, ''))
              ? 'Double'
              : /^&[Hh]/.test(valueToken.text)
                ? 'Long'
                : 'Long';
        } else if (lower(valueToken.text) === 'true' || lower(valueToken.text) === 'false') {
          typeName = 'Boolean';
        }
      }
    }

    const prefixLabel = options.constant
      ? options.moduleLevel ? `${accessLabel(options.access)} Const` : 'Const'
      : options.moduleLevel
        ? `${accessLabel(options.access)}${options.withEvents ? ' WithEvents' : ''}`
        : options.isStatic
          ? 'Static'
          : 'Dim';
    const signature = `${prefixLabel} ${name}${array ? '()' : ''}${typeName !== undefined ? ` As ${typeName}` : ''}`;

    const symbol: VbaSymbol = {
      id: symbolId(context, name, options.idScope, options.constant ? 'constant' : 'variable'),
      name,
      kind: options.constant ? 'constant' : 'variable',
      uri: context.uri,
      range: range(declaratorStart.line, declaratorStart.start, declaratorEnd.line, declaratorEnd.end),
      selectionRange: range(nameToken.line, nameToken.start, nameToken.line, nameToken.start + name.length),
      ...(options.scope === undefined ? {} : { scope: options.scope }),
      ...(options.container === undefined ? {} : { container: options.container }),
      ...(typeName === undefined ? {} : { typeName }),
      access: options.access,
      signature,
    };
    addSymbol(context, symbol);
    if (options.withEvents) {
      context.withEventsNames.add(name.toLowerCase());
    }
  }
}

function lookupTypeAfterAs(part: VbaToken[], startIndex: number): string | undefined {
  for (let i = startIndex; i < part.length; i += 1) {
    const token = part[i];
    if (token === undefined) {
      continue;
    }
    if (lower(token.text) === 'new') {
      continue;
    }
    if (token.kind === 'identifier' || token.kind === 'keyword') {
      let name = token.text;
      let next = i + 1;
      while (part[next]?.text === '.' && ['identifier', 'keyword'].includes(part[next + 1]?.kind ?? '')) {
        name += `.${part[next + 1]!.text}`;
        next += 2;
      }
      if (part[next]?.text === '(' && part[next + 1]?.text === ')') { name += '()'; }
      return name;
    }
    break;
  }
  return undefined;
}

function parseProcedure(
  context: ParserContext,
  tokens: VbaToken[],
  index: number,
  prefix: DeclPrefix,
  line: number,
  endChar: number,
): void {
  const first = lower(tokens[index]?.text);
  let external = false;
  let ptrSafe = false;
  let procKeyword = first;
  let nameIndex = index + 1;
  let accessor: string | undefined;
  let libName: string | undefined;
  let aliasName: string | undefined;

  if (first === 'declare') {
    external = true;
    let cursor = index + 1;
    if (lower(tokens[cursor]?.text) === 'ptrsafe') {
      ptrSafe = true;
      cursor += 1;
    }
    procKeyword = lower(tokens[cursor]?.text);
    if (procKeyword !== 'sub' && procKeyword !== 'function') {
      context.diagnostics.push(
        makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.malformedDeclare', 'Declare 后需要 Sub 或 Function。'),
      );
      return;
    }
    nameIndex = cursor + 1;
  } else if (first === 'property') {
    accessor = lower(tokens[index + 1]?.text);
    if (accessor !== 'get' && accessor !== 'let' && accessor !== 'set') {
      context.diagnostics.push(
        makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.malformedProperty', 'Property 后需要 Get、Let 或 Set。'),
      );
      return;
    }
    nameIndex = index + 2;
  }

  const nameToken = tokens[nameIndex];
  if (nameToken === undefined || nameToken.kind !== 'identifier') {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.missingName', '过程声明缺少名称。'),
    );
    return;
  }

  if (external) {
    for (let i = nameIndex; i < tokens.length; i += 1) {
      const word = lower(tokens[i]?.text);
      if (word === 'lib') {
        const libToken = tokens[i + 1];
        if (libToken?.kind === 'string') {
          libName = libToken.text.slice(1, -1);
        }
      } else if (word === 'alias') {
        const aliasToken = tokens[i + 1];
        if (aliasToken?.kind === 'string') {
          aliasName = aliasToken.text.slice(1, -1);
        }
      }
    }
  }

  let openIndex = -1;
  for (let i = nameIndex + 1; i < tokens.length; i += 1) {
    const token = tokens[i];
    if (token?.text === '(') {
      openIndex = i;
      break;
    }
  }
  const parameterResult = openIndex >= 0
    ? parseParameters(tokens, openIndex)
    : { parameters: [] as ParsedParameter[], closeIndex: nameIndex };
  const parameters = parameterResult.parameters;

  let returnType: string | undefined;
  if (procKeyword === 'function' || (first === 'property' && accessor === 'get')) {
    const searchFrom = Math.max(parameterResult.closeIndex, nameIndex) + 1;
    for (let i = searchFrom; i < tokens.length; i += 1) {
      if (lower(tokens[i]?.text) === 'as') {
        let typeIndex = -1;
        for (let j = i + 1; j < tokens.length; j += 1) {
          const token = tokens[j];
          if (token === undefined) {
            continue;
          }
          if (lower(token.text) === 'new') {
            continue;
          }
          if (token.kind === 'identifier' || token.kind === 'keyword') {
            typeIndex = j;
          }
          break;
        }
        if (typeIndex >= 0) {
          returnType = lookupTypeAfterAs(tokens, typeIndex);
        }
        break;
      }
    }
  }

  const access: Access = prefix.access ?? 'public';
  const accessWord = accessLabel(access);
  const paramsText = parameters.map(parameterSignature).join(', ');
  const kindPart = first === 'property' ? `property:${accessor ?? ''}` : procKeyword;
  let signature: string;
  if (external) {
    signature = `${accessWord} Declare${ptrSafe ? ' PtrSafe' : ''} ${procKeyword === 'function' ? 'Function' : 'Sub'} ${nameToken.text}`
      + (libName !== undefined ? ` Lib "${libName}"` : '')
      + (aliasName !== undefined ? ` Alias "${aliasName}"` : '')
      + ` (${paramsText})`
      + (returnType !== undefined ? ` As ${returnType}` : '');
  } else if (first === 'property') {
    signature = `${accessWord} Property ${accessor === undefined ? '' : `${accessor.slice(0, 1).toUpperCase()}${accessor.slice(1)}`} ${nameToken.text}(${paramsText})`
      + (returnType !== undefined ? ` As ${returnType}` : '');
  } else {
    signature = `${accessWord} ${procKeyword === 'function' ? 'Function' : 'Sub'} ${nameToken.text}(${paramsText})`
      + (returnType !== undefined ? ` As ${returnType}` : '');
  }

  closeUnclosedBeforeProcedure(context, line, endChar);
  const moduleLevel = !context.stack.some((entry) => entry.type === 'proc');
  const enclosingProc = context.currentProc;
  const idScope = first === 'property'
    ? procScopeKey(context, nameToken.text, accessor)
    : moduleLevel
      ? moduleScope(context)
      : currentProcedureScopeKey(context);
  const scopeField = moduleLevel ? undefined : enclosingProc?.symbol?.id;
  const container = moduleLevel ? context.moduleId : enclosingProc?.symbol?.id;
  const documentation = collectDocumentation(context, line);
  const symbol: VbaSymbol = {
    id: symbolId(context, nameToken.text, idScope, kindPart),
    name: nameToken.text,
    kind: first === 'property' ? 'property' : procKeyword === 'function' ? 'function' : 'sub',
    uri: context.uri,
    range: range(line, tokens[index]?.start ?? 0, line, endChar),
    selectionRange: tokenRange(nameToken),
    ...(scopeField === undefined ? {} : { scope: scopeField }),
    ...(container === undefined ? {} : { container }),
    ...(returnType === undefined ? {} : { typeName: returnType }),
    parameters: publicParameters(parameters),
    access,
    signature,
    ...(documentation === undefined ? {} : { documentation }),
    ...(external ? { external: true } : {}),
  };

  addSymbol(context, symbol);
  for (const parameter of parameters) {
    addSymbol(context, {
      id: symbolId(context, parameter.name, idScope, 'parameter'),
      name: parameter.name,
      kind: 'parameter',
      uri: context.uri,
      range: tokenRange(parameter.token),
      selectionRange: tokenRange(parameter.token),
      scope: symbol.id,
      container: symbol.id,
      typeName: parameter.typeName,
      access: 'private',
      signature: parameterSignature(parameter),
    });
  }
  if (external) {
    return;
  }
  const entry: BlockEntry = {
    type: 'proc',
    startLine: line,
    startChar: tokens[index]?.start ?? 0,
    symbol,
    procKeyword,
    accessor,
    name: nameToken.text,
  };
  context.stack.push(entry);
  context.currentProc = entry;
}

function parseEvent(context: ParserContext, tokens: VbaToken[], index: number, line: number, endChar: number): void {
  const nameToken = tokens[index + 1];
  if (nameToken === undefined || nameToken.kind !== 'identifier') {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.missingName', 'Event 声明缺少名称。'),
    );
    return;
  }
  let openIndex = -1;
  for (let i = index + 2; i < tokens.length; i += 1) {
    if (tokens[i]?.text === '(') {
      openIndex = i;
      break;
    }
  }
  const parameters = openIndex >= 0 ? parseParameters(tokens, openIndex).parameters : [];
  const access: Access = 'public';
  const symbol: VbaSymbol = {
    id: symbolId(context, nameToken.text, moduleScope(context), 'event'),
    name: nameToken.text,
    kind: 'event',
    uri: context.uri,
    range: range(line, tokens[index]?.start ?? 0, line, endChar),
    selectionRange: tokenRange(nameToken),
    container: context.moduleId,
    parameters: publicParameters(parameters),
    access,
    signature: `${accessLabel(access)} Event ${nameToken.text}(${parameters.map(parameterSignature).join(', ')})`,
  };
  addSymbol(context, symbol);
  for (const parameter of parameters) {
    addSymbol(context, {
      id: symbolId(context, parameter.name, `${context.moduleName}.${nameToken.text}`, 'parameter'),
      name: parameter.name,
      kind: 'parameter',
      uri: context.uri,
      range: tokenRange(parameter.token),
      selectionRange: tokenRange(parameter.token),
      scope: symbol.id,
      container: symbol.id,
      typeName: parameter.typeName,
      access: 'private',
      signature: parameterSignature(parameter),
    });
  }
}

function parseTypeOrEnum(
  context: ParserContext,
  tokens: VbaToken[],
  index: number,
  prefix: DeclPrefix,
  line: number,
  endChar: number,
  type: 'type' | 'enum',
): void {
  if (context.stack.some((entry) => entry.type === 'proc')) {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.nestedType', `${type === 'type' ? 'Type' : 'Enum'} 声明只能位于模块级。`),
    );
    return;
  }
  const nameToken = tokens[index + 1];
  if (nameToken === undefined || nameToken.kind !== 'identifier') {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[index]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.missingName', `${type === 'type' ? 'Type' : 'Enum'} 声明缺少名称。`),
    );
    return;
  }
  while (context.stack.length > 0 && (context.stack[context.stack.length - 1]?.type === 'type' || context.stack[context.stack.length - 1]?.type === 'enum')) {
    const entry = context.stack[context.stack.length - 1]!;
    reportMissingEnd(context, entry);
    popBlock(context, line, endChar);
  }
  const access = prefix.access ?? 'public';
  const symbol: VbaSymbol = {
    id: symbolId(context, nameToken.text, moduleScope(context), type),
    name: nameToken.text,
    kind: type,
    uri: context.uri,
    range: range(line, tokens[index]?.start ?? 0, line, endChar),
    selectionRange: tokenRange(nameToken),
    container: context.moduleId,
    access,
    signature: `${accessLabel(access)} ${type === 'type' ? 'Type' : 'Enum'} ${nameToken.text}`,
  };
  addSymbol(context, symbol);
  context.stack.push({
    type,
    startLine: line,
    startChar: tokens[index]?.start ?? 0,
    symbol,
    name: nameToken.text,
  });
}

function parseField(
  context: ParserContext,
  tokens: VbaToken[],
  container: BlockEntry,
  line: number,
): void {
  if (container.symbol === undefined || container.name === undefined) {
    return;
  }
  const parts = container.type === 'type' ? splitTopLevel(tokens, ',') : [tokens];
  for (const part of parts) {
    if (part.length === 0) {
      continue;
    }
    let nameToken: VbaToken | undefined;
    for (const token of part) {
      if (token.kind === 'identifier') {
        nameToken = token;
        break;
      }
    }
    if (nameToken === undefined) {
      continue;
    }
    const last = part[part.length - 1] ?? nameToken;
    let typeName: string | undefined;
    let asIndex = -1;
    for (let i = 0; i < part.length; i += 1) {
      if (lower(part[i]?.text) === 'as') {
        asIndex = i;
        typeName = lookupTypeAfterAs(part, i + 1);
        break;
      }
    }
    let array = false;
    let sawName = false;
    for (let i = 0; i < part.length; i += 1) {
      const token = part[i];
      if (token === nameToken) {
        sawName = true;
        continue;
      }
      if (sawName && token?.text === '(') {
        array = true;
        break;
      }
      if (asIndex >= 0 && i >= asIndex) {
        break;
      }
    }
    const signature = container.type === 'type'
      ? `${nameToken.text}${array ? '()' : ''} As ${typeName ?? 'Variant'}`
      : nameToken.text;
    const symbol: VbaSymbol = {
      id: symbolId(context, nameToken.text, `${context.moduleName}.${container.name}`, container.type === 'type' ? 'field' : 'enumMember'),
      name: nameToken.text,
      kind: container.type === 'type' ? 'field' : 'enumMember',
      uri: context.uri,
      range: range(part[0]?.line ?? line, part[0]?.start ?? nameToken.start, last.line, last.end),
      selectionRange: tokenRange(nameToken),
      scope: container.symbol.id,
      container: container.name,
      ...(typeName === undefined ? {} : { typeName }),
      access: 'public',
      signature,
    };
    addSymbol(context, symbol);
  }
}

/** Option 语句：必须在任何声明之前，且同一 Option 不得重复（PRE-01）。 */
function handleOptionStatement(context: ParserContext, tokens: VbaToken[], line: number, endChar: number): void {
  const name = lower(tokens[1]?.text) ?? '';
  if (context.declarationSeen) {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.optionPosition', 'Option 语句必须位于所有声明之前。'),
    );
  }
  if (name !== '') {
    if (context.optionsSeen.has(name)) {
      context.diagnostics.push(
        makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.optionDuplicate', `Option ${name} 重复。`),
      );
    } else {
      context.optionsSeen.add(name);
    }
  }
  if (name === 'explicit') {
    context.explicit = true;
  }
}

/** Deftype 语句必须位于所有声明之前（PRE-09）。 */
function handleDeftypeStatement(context: ParserContext, tokens: VbaToken[], line: number, endChar: number): void {
  if (context.declarationSeen) {
    context.diagnostics.push(
      makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.deftypePosition', 'Deftype 语句必须位于所有声明之前。'),
    );
  }
}

function handleDeclarations(
  context: ParserContext,
  tokens: VbaToken[],
  line: number,
  endLine: number,
  endChar: number,
): void {
  // 方括号转义标识符开头的语句（`[Dim] = 1`）不是声明，交回普通语句处理（LEX-07）。
  if (tokens[0]?.escaped === true) {
    return;
  }
  const prefix = parseDeclPrefix(tokens);
  const index = prefix.index;
  const word = lower(tokens[index]?.text);

  // 非 Option 的声明与语句都会让后续 Option/Deftype 失去合法位置（PRE-01/09）。
  context.declarationSeen = true;

  if (word === 'implements') {
    return;
  }
  if (word === 'declare') {
    parseProcedure(context, tokens, index, prefix, line, endChar);
    return;
  }
  if (word === 'sub' || word === 'function' || word === 'property') {
    parseProcedure(context, tokens, index, prefix, line, endChar);
    return;
  }
  if (word === 'event') {
    parseEvent(context, tokens, index, line, endChar);
    return;
  }
  if (word === 'type' || word === 'enum') {
    parseTypeOrEnum(context, tokens, index, prefix, line, endChar, word);
    return;
  }

  const moduleLevel = !context.stack.some((entry) => entry.type === 'proc');
  if (word === 'dim' || word === 'const') {
    const access = moduleLevel
      ? word === 'const'
        ? accessFrom(prefix, true)
        : prefix.access ?? 'private'
      : 'private';
    parseVariableDeclaration(context, tokens, index + 1, {
      access,
      moduleLevel,
      isStatic: prefix.isStatic,
      withEvents: prefix.withEvents,
      constant: word === 'const',
      idScope: moduleLevel ? moduleScope(context) : currentProcedureScopeKey(context),
      ...(moduleLevel ? {} : { scope: context.currentProc?.symbol?.id }),
      container: moduleLevel ? context.moduleId : context.currentProc?.symbol?.id,
    });
    return;
  }

  const current = tokens[index];
  if (current !== undefined && (current.kind === 'identifier' || lower(current.text) === 'withevents') &&
      (prefix.access !== undefined || prefix.global || prefix.isStatic || prefix.withEvents)) {
    parseVariableDeclaration(context, tokens, index, {
      access: accessFrom(prefix, moduleLevel),
      moduleLevel,
      isStatic: prefix.isStatic,
      withEvents: prefix.withEvents,
      constant: false,
      idScope: moduleLevel ? moduleScope(context) : currentProcedureScopeKey(context),
      ...(moduleLevel ? {} : { scope: context.currentProc?.symbol?.id }),
      container: moduleLevel ? context.moduleId : context.currentProc?.symbol?.id,
    });
  }
}

function splitStatements(group: LogicalLine): Array<{ tokens: VbaToken[]; labelOnly: boolean; line: number; terminatedByColon: boolean }> {
  const result: Array<{ tokens: VbaToken[]; labelOnly: boolean; line: number; terminatedByColon: boolean }> = [];
  let current: VbaToken[] = [];
  let depth = 0;
  let labelCandidate = true;
  const push = (terminatedByColon: boolean): void => {
    if (current.length === 0) {
      return;
    }
    const labelOnly = terminatedByColon && labelCandidate && current.length === 1 && current[0]?.kind === 'identifier';
    result.push({ tokens: current, labelOnly, line: current[0]?.line ?? group.startLine, terminatedByColon });
    current = [];
    labelCandidate = false;
  };
  for (const token of group.tokens) {
    if (token.text === '(') {
      depth += 1;
    } else if (token.text === ')') {
      depth = Math.max(0, depth - 1);
    }
    if (token.text === ':' && depth === 0) {
      push(true);
      labelCandidate = true;
      continue;
    }
    current.push(token);
  }
  push(false);
  return result;
}

function processBlockStatement(
  context: ParserContext,
  tokens: VbaToken[],
  line: number,
  endLine: number,
  endChar: number,
): boolean {
  const first = lower(tokens[0]?.text);
  const second = lower(tokens[1]?.text);

  if (first === 'end' && second !== '') {
    if (second === 'sub') {
      closeProcedure(context, 'sub', line, endChar);
      return true;
    }
    if (second === 'function') {
      closeProcedure(context, 'function', line, endChar);
      return true;
    }
    if (second === 'property') {
      closeProcedure(context, 'property', line, endChar);
      return true;
    }
    if (second === 'if') {
      closeByType(context, 'if', line, endChar);
      return true;
    }
    if (second === 'with') {
      closeByType(context, 'with', line, endChar);
      return true;
    }
    if (second === 'select') {
      closeByType(context, 'select', line, endChar);
      return true;
    }
    if (second === 'type') {
      closeByType(context, 'type', line, endChar);
      return true;
    }
    if (second === 'enum') {
      closeByType(context, 'enum', line, endChar);
      return true;
    }
  }
  if (first === 'next') {
    // `Next a, b` 用一条语句关闭多层 For（VBA-10）。逗号须在括号外。
    let commas = 0;
    let depth = 0;
    for (let i = 1; i < tokens.length; i += 1) {
      const text = tokens[i]?.text;
      if (text === '(') { depth += 1; continue; }
      if (text === ')') { depth = Math.max(0, depth - 1); continue; }
      if (text === ',' && depth === 0) { commas += 1; }
    }
    for (let i = 0; i <= commas; i += 1) { closeByType(context, 'for', line, endChar); }
    return true;
  }
  if (first === 'loop') {
    closeByType(context, 'do', line, endChar);
    return true;
  }
  if (first === 'wend') {
    closeByType(context, 'while', line, endChar);
    return true;
  }
  if (first === 'else' || first === 'elseif') {
    const top = context.stack[context.stack.length - 1];
    if (top === undefined || top.type !== 'if') {
      context.diagnostics.push(
        makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.unmatchedElse', `${first === 'else' ? 'Else' : 'ElseIf'} 没有对应的 If。`),
      );
    }
    return true;
  }
  return false;
}

function openBlockStatements(
  context: ParserContext,
  tokens: VbaToken[],
  line: number,
  endChar: number,
  terminatedByColon: boolean,
): boolean {
  const first = lower(tokens[0]?.text);
  const second = lower(tokens[1]?.text);
  if (first === 'if') {
    let thenIndex = -1;
    let depth = 0;
    for (let i = 1; i < tokens.length; i += 1) {
      const text = tokens[i]?.text;
      if (text === '(') {
        depth += 1;
      } else if (text === ')') {
        depth = Math.max(0, depth - 1);
      }
      if (depth === 0 && lower(text) === 'then') {
        thenIndex = i;
        break;
      }
    }
    if (thenIndex < 0) {
      context.diagnostics.push(
        makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.missingThen', 'If 缺少 Then。'),
      );
      return true;
    }
    // 行内冒号表示 Then 后的语句从同一逻辑行的下一个语句段开始；
    // 这是单行 If，不是块 If。
    if (thenIndex < tokens.length - 1 || terminatedByColon) {
      return true;
    }
    context.stack.push({ type: 'if', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  if (first === 'with') {
    context.stack.push({ type: 'with', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  if (first === 'for') {
    context.stack.push({ type: 'for', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  if (first === 'do') {
    context.stack.push({ type: 'do', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  if (first === 'while') {
    context.stack.push({ type: 'while', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  if (first === 'select' && second === 'case') {
    context.stack.push({ type: 'select', startLine: line, startChar: tokens[0]?.start ?? 0 });
    return true;
  }
  return false;
}

function startsDeclarationStatement(tokens: VbaToken[]): boolean {
  const prefix = parseDeclPrefix(tokens);
  const word = lower(tokens[prefix.index]?.text);
  return word === 'sub' || word === 'function' || word === 'property' || word === 'declare' ||
    word === 'type' || word === 'enum' || word === 'event';
}

function isSingleLineIf(tokens: VbaToken[], terminatedByColon = false): boolean {
  if (lower(tokens[0]?.text) !== 'if') {
    return false;
  }
  let depth = 0;
  for (let i = 1; i < tokens.length; i += 1) {
    const text = tokens[i]?.text;
    if (text === '(') {
      depth += 1;
    } else if (text === ')') {
      depth = Math.max(0, depth - 1);
    }
    if (depth === 0 && lower(text) === 'then') {
      return i < tokens.length - 1 || terminatedByColon;
    }
  }
  return false;
}

function stripLineNumber(tokens: VbaToken[]): VbaToken[] {
  const first = tokens[0];
  if (first !== undefined && first.kind === 'number' && tokens.length > 1 && !first.text.startsWith('#')) {
    return tokens.slice(1);
  }
  return tokens;
}

function parseCode(context: ParserContext, groups: LogicalLine[]): void {
  for (const group of groups) {
    if (!context.lineActive[group.startLine] || context.protectedLines.has(group.startLine) || isDirective(group)) {
      continue;
    }
    const statements = splitStatements(group);
    let inlineRemainder = false;
    for (const statement of statements) {
      if (statement.labelOnly) {
        continue;
      }
      const tokens = stripLineNumber(statement.tokens);
      if (tokens.length === 0 || inlineRemainder) {
        continue;
      }
      const line = tokens[0]?.line ?? group.startLine;
      const endLine = tokens[tokens.length - 1]?.line ?? group.endLine;
      const endChar = tokens[tokens.length - 1]?.end ?? 0;
      const firstWord = lower(tokens[0]?.text);

      if (firstWord === 'option') {
        handleOptionStatement(context, tokens, line, endChar);
        continue;
      }
      if (firstWord.startsWith('def') && KEYWORDS.has(firstWord)) {
        handleDeftypeStatement(context, tokens, line, endChar);
        continue;
      }
      if (firstWord === 'exit' && lower(tokens[1]?.text) === 'while') {
        // VBA 没有 Exit While（STMT-08b）。
        context.diagnostics.push(
          makeDiagnostic(line, tokens[0]?.start ?? 0, line, endChar, SEVERITY_ERROR, 'syntax.exitWhile', 'VBA 不存在 Exit While 语句。'),
        );
        continue;
      }

      if (processBlockStatement(context, tokens, line, endLine, endChar)) {
        if (isSingleLineIf(tokens, statement.terminatedByColon)) {
          inlineRemainder = true;
        }
        continue;
      }

      const top = context.stack[context.stack.length - 1];
      if (top !== undefined && (top.type === 'type' || top.type === 'enum')) {
        if (!startsDeclarationStatement(tokens)) {
          parseField(context, tokens, top, line);
          continue;
        }
        while (context.stack.length > 0) {
          const openType = context.stack[context.stack.length - 1];
          if (openType === undefined || (openType.type !== 'type' && openType.type !== 'enum')) {
            break;
          }
          const previous = previousLineEnd(context, line);
          reportMissingEnd(context, openType);
          popBlock(context, previous.line, previous.end, false);
        }
      }

      if (openBlockStatements(context, tokens, line, endChar, statement.terminatedByColon)) {
        context.declarationSeen = true;
        if (isSingleLineIf(tokens, statement.terminatedByColon)) {
          inlineRemainder = true;
        }
        continue;
      }

      context.declarationSeen = true;
      handleDeclarations(context, tokens, line, endLine, endChar);
    }
  }

  const lastLine = Math.max(0, context.lines.length - 1);
  const lastChar = (context.lines[lastLine] ?? '').length;
  while (context.stack.length > 0) {
    const entry = context.stack[context.stack.length - 1]!;
    reportMissingEnd(context, entry);
    popBlock(context, lastLine, lastChar, false);
  }
}

function detectKindFromUri(uri: string): ModuleKind | undefined {
  let decoded = uri;
  try {
    decoded = decodeURIComponent(uri);
  } catch {
    decoded = uri;
  }
  const normalized = decoded.replace(/\\/g, '/');
  if (/\/表格对象\//.test(normalized)) {
    return 'document';
  }
  if (/\/类模块\//.test(normalized)) {
    return 'class';
  }
  if (/\/窗体\//.test(normalized)) {
    return 'form';
  }
  if (/\/模块\//.test(normalized)) {
    return 'standard';
  }
  return undefined;
}

function detectKindFromExtension(uri: string): ModuleKind | undefined {
  const lowerUri = uri.toLowerCase();
  if (lowerUri.endsWith('.bas')) {
    return 'standard';
  }
  if (lowerUri.endsWith('.cls')) {
    return 'class';
  }
  if (lowerUri.endsWith('.frm')) {
    return 'form';
  }
  return undefined;
}

function baseNameFromUri(uri: string): string {
  let decoded = uri;
  try {
    decoded = decodeURIComponent(uri);
  } catch {
    decoded = uri;
  }
  const normalized = decoded.replace(/\\/g, '/');
  const last = normalized.slice(normalized.lastIndexOf('/') + 1);
  return last.replace(/\.(bas|cls|frm)$/i, '');
}

function sortDiagnostics(diagnostics: Diagnostic[]): Diagnostic[] {
  diagnostics.sort((left, right) => {
    if (left.range.start.line !== right.range.start.line) {
      return left.range.start.line - right.range.start.line;
    }
    if (left.range.start.character !== right.range.start.character) {
      return left.range.start.character - right.range.start.character;
    }
    return String(left.code ?? '').localeCompare(String(right.code ?? ''));
  });
  return diagnostics;
}

function sortAndDedupeFolds(folds: FoldingRange[]): FoldingRange[] {
  const seen = new Set<string>();
  const result: FoldingRange[] = [];
  for (const fold of folds) {
    const key = `${fold.startLine}:${fold.endLine}:${fold.kind ?? ''}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(fold);
  }
  result.sort((left, right) => left.startLine - right.startLine || left.endLine - right.endLine);
  return result;
}

const EVENT_PREFIXES: Set<string> = new Set([
  'worksheet', 'workbook', 'chart', 'sheet', 'document', 'thisdocument',
  'userform', 'form', 'report', 'class', 'application',
  'commandbutton', 'textbox', 'combobox', 'listbox', 'checkbox', 'optionbutton',
  'togglebutton', 'spinbutton', 'scrollbar', 'label', 'image', 'frame',
  'multipage', 'tabstrip', 'treeview', 'listview', 'refedit', 'oleobject',
  'webbrowser', 'commandbarbutton', 'commandbarcombobox', 'commandbarcontrol',
]);

function markEventHandlers(context: ParserContext, kind: ModuleKind): void {
  const knownPrefixes = new Set(EVENT_PREFIXES);
  knownPrefixes.add(context.moduleName.toLowerCase());
  const withEvents = [...context.withEventsNames];
  for (const symbol of context.symbols) {
    if (symbol.kind !== 'sub' || !symbol.name.includes('_')) {
      continue;
    }
    const lowerName = symbol.name.toLowerCase();
    const firstPrefix = lowerName.slice(0, lowerName.indexOf('_'));
    const withEventsMatch = withEvents.some((prefix) => lowerName.startsWith(`${prefix}_`));
    if (withEventsMatch ||
        ((kind === 'document' || kind === 'form') && knownPrefixes.has(firstPrefix))) {
      symbol.eventHandler = true;
    }
  }
}

/**
 * 解析一个 VBA 模块。词法与结构容错；不执行任何源码，条件编译只做受限表达式求值。
 */
export function parseModule(uri: string, text: string, options: ParseOptions = {}): ParsedModule {
  const lines = text.split(/\r\n|\n|\r/);
  const header = scanHeader(lines);
  const groups = buildLogicalLines(lines);

  const conditional = evaluateConditionals(groups, lines.length, header.protectedLines, options);
  const lineActive = conditional.lineActive;

  const lexDiagnostics: Diagnostic[] = [];
  for (let i = 0; i < lines.length; i += 1) {
    if (!lineActive[i] || header.protectedLines.has(i)) {
      continue;
    }
    lexDiagnostics.push(...lexLine(lines[i] ?? '', i).diagnostics);
  }

  // 逻辑行最多由 25 个物理行组成（24 个连接符）；第 25 个连接符报错（LEX-02）。
  for (const group of groups) {
    let count = 0;
    let reported = false;
    for (let line = group.startLine; line <= group.endLine && !reported; line += 1) {
      if (!lineActive[line] || header.protectedLines.has(line)) { continue; }
      if ((lines[line] ?? '').trimEnd().endsWith('_')) {
        count += 1;
        if (count > 24) {
          reported = true;
          lexDiagnostics.push(
            makeDiagnostic(line, Math.max(0, (lines[line] ?? '').indexOf('_')), line, (lines[line] ?? '').length, SEVERITY_ERROR, 'lex.tooManyContinuations', '行连接超过 24 个：逻辑行最多由 25 个物理行组成。'),
          );
        }
      }
    }
  }

  const headerKind: ModuleKind | undefined = header.versionForm
    ? 'form'
    : header.versionClass
      ? header.predeclaredId === true ? 'document' : 'class'
      : undefined;
  const kind: ModuleKind = options.moduleKind
    ?? headerKind
    ?? detectKindFromUri(uri)
    ?? detectKindFromExtension(uri)
    ?? 'standard';

  const attributeName = header.attributes.get('vb_name');
  const baseName = baseNameFromUri(uri);
  const name = options.moduleName ?? attributeName ?? (baseName !== '' ? baseName : 'Module');
  const moduleId = `${uri}#${name}#module`;

  const moduleKindLabel = kind === 'standard' ? 'Module' : kind === 'class' ? 'Class' : kind === 'document' ? 'Document' : 'Form';
  const moduleSymbol: VbaSymbol = {
    id: moduleId,
    name,
    kind: kind === 'standard' ? 'module' : 'class',
    uri,
    range: range(0, 0, Math.max(0, lines.length - 1), (lines[lines.length - 1] ?? '').length),
    selectionRange: header.nameLocation === undefined
      ? range(0, 0, 0, 0)
      : range(header.nameLocation.line, header.nameLocation.start, header.nameLocation.line, header.nameLocation.end),
    access: 'public',
    signature: `${moduleKindLabel} ${name}`,
  };

  const context: ParserContext = {
    uri,
    moduleName: name,
    moduleId,
    lines,
    lineActive,
    symbols: [moduleSymbol],
    diagnostics: [...conditional.diagnostics, ...lexDiagnostics],
    folds: [],
    stack: [],
    explicit: false,
    optionsSeen: new Set<string>(),
    declarationSeen: false,
    withEventsNames: new Set<string>(),
    protectedLines: header.protectedLines,
  };

  parseCode(context, groups);
  markEventHandlers(context, kind);

  for (const [procName, description] of header.descriptions) {
    for (const symbol of context.symbols) {
      if (symbol.name.toLowerCase() === procName && symbol.documentation === undefined &&
          (symbol.kind === 'sub' || symbol.kind === 'function' || symbol.kind === 'property')) {
        symbol.documentation = description;
      }
    }
  }

  const commentBlocks: Array<{ start: number; end: number }> = [];
  let blockStart = -1;
  for (let i = 0; i < lines.length; i += 1) {
    const trimmed = (lines[i] ?? '').trim();
    const isComment = !header.protectedLines.has(i) && lineActive[i] === true &&
      (trimmed.startsWith("'") || /^rem(?![A-Za-z0-9_])/i.test(trimmed));
    if (isComment) {
      if (blockStart < 0) {
        blockStart = i;
      }
    } else if (blockStart >= 0) {
      if (i - 1 > blockStart) {
        commentBlocks.push({ start: blockStart, end: i - 1 });
      }
      blockStart = -1;
    }
  }
  if (blockStart >= 0 && lines.length - 1 > blockStart) {
    commentBlocks.push({ start: blockStart, end: lines.length - 1 });
  }
  for (const commentBlock of commentBlocks) {
    addFold(context, commentBlock.start, commentBlock.end, 'comment');
  }

  if (header.designStart !== undefined && header.designEnd !== undefined && header.designEnd > header.designStart) {
    addFold(context, header.designStart, header.designEnd, 'region');
  }

  return {
    uri,
    text,
    lines,
    name,
    kind,
    explicit: context.explicit,
    tokens: tokenize(text),
    symbols: context.symbols,
    diagnostics: sortDiagnostics(context.diagnostics),
    folds: sortAndDedupeFolds(context.folds),
    protectedLines: new Set(header.protectedLines),
    inactiveLines: new Set(lineActive.map((active, index) => (active ? -1 : index)).filter((index) => index >= 0)),
    unknownConditional: conditional.unknownConditional,
  };
}
