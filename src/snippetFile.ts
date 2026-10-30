/**
 * 自定义代码片段文件的解析（不依赖 vscode 模块）。
 *
 * 片段文件沿用 VS Code 的 snippet JSON 结构：
 * { "名称": { "prefix": "sub" | ["sub","subr"], "body": ["第 1 行", "第 2 行"], "description": "..." } }
 * 为方便手工维护，允许 // 与 /* *\/ 注释。
 */

export interface CustomSnippet {
  /** 片段条目名（JSON 键）。 */
  name: string;
  /** 触发前缀，可多个。 */
  prefixes: string[];
  /** 已用换行连接好的片段正文（保留 $1、${1:默认值} 等 VS Code 片段语法）。 */
  body: string;
  description?: string;
}

export class SnippetFileError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'SnippetFileError';
  }
}

/** 去掉 JSONC 注释，字符串内的 // 与 /* 保持原样。 */
export function stripJsonComments(text: string): string {
  let output = '';
  let index = 0;
  let inString = false;
  while (index < text.length) {
    const char = text[index]!;
    const next = text[index + 1];
    if (inString) {
      output += char;
      if (char === '\\' && next !== undefined) {
        output += next;
        index += 2;
        continue;
      }
      if (char === '"') { inString = false; }
      index += 1;
      continue;
    }
    if (char === '"') {
      inString = true;
      output += char;
      index += 1;
      continue;
    }
    if (char === '/' && next === '/') {
      while (index < text.length && text[index] !== '\n') { index += 1; }
      continue;
    }
    if (char === '/' && next === '*') {
      index += 2;
      while (index < text.length && !(text[index] === '*' && text[index + 1] === '/')) { index += 1; }
      index += 2;
      continue;
    }
    output += char;
    index += 1;
  }
  return output;
}

function readPrefixes(value: unknown, name: string): string[] {
  if (typeof value === 'string' && value.trim()) { return [value.trim()]; }
  if (Array.isArray(value)) {
    const list = value.filter((item): item is string => typeof item === 'string' && item.trim().length > 0).map(item => item.trim());
    if (list.length) { return list; }
  }
  throw new SnippetFileError(`片段「${name}」缺少有效的 prefix。`);
}

function readBody(value: unknown, name: string): string {
  if (typeof value === 'string' && value.length > 0) { return value; }
  if (Array.isArray(value)) {
    const lines = value.filter((item): item is string => typeof item === 'string');
    if (lines.length) { return lines.join('\n'); }
  }
  throw new SnippetFileError(`片段「${name}」缺少有效的 body。`);
}

export function parseCustomSnippets(text: string): CustomSnippet[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(stripJsonComments(text));
  } catch (error) {
    throw new SnippetFileError(error instanceof Error ? error.message : 'JSON 解析失败。');
  }
  if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new SnippetFileError('片段文件的根节点必须是对象。');
  }
  const snippets: CustomSnippet[] = [];
  for (const [name, value] of Object.entries(parsed as Record<string, unknown>)) {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) { continue; }
    const entry = value as Record<string, unknown>;
    const prefixes = readPrefixes(entry['prefix'], name);
    const body = readBody(entry['body'], name);
    const description = typeof entry['description'] === 'string' ? entry['description'] : undefined;
    snippets.push({ name, prefixes, body, description });
  }
  return snippets;
}

/** 按当前输入前缀筛选片段；前缀不区分大小写。 */
export function matchSnippets(snippets: readonly CustomSnippet[], word: string): CustomSnippet[] {
  const prefix = word.trim().toLowerCase();
  if (!prefix) { return []; }
  return snippets.filter(snippet => snippet.prefixes.some(item => item.toLowerCase().startsWith(prefix)));
}

/** 新建片段文件时的模板（含注释，说明字段含义）。 */
export const CUSTOM_SNIPPET_TEMPLATE = `{
  // VBAdminx 自定义代码片段。保存后立即生效。
  // prefix：触发前缀，可写字符串或字符串数组；body：插入内容，每项一行。
  "Log message": {
    "prefix": ["log", "debuglog"],
    "body": [
      "Debug.Print \\"\${1:message}\\""
    ],
    "description": "Print a message to the Immediate window"
  }
}
`;
