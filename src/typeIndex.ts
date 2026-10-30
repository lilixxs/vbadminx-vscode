import type { PortableTypeIndex, VbaParameter } from './language/model';

/**
 * 可移植类型索引的校验与序列化。本模块不得依赖 vscode，便于在扩展主机与普通 Node 测试中使用。
 * 契约见 src/language/model.ts 的 PortableTypeIndex 与 schemas/type-index.schema.json。
 */

export const TYPE_INDEX_FORMAT_VERSION = 1;
export const TYPE_INDEX_KIND_VALUES = ['class', 'interface', 'enum', 'module'] as const;
export const TYPE_INDEX_MEMBER_KIND_VALUES = ['method', 'property', 'constant', 'event'] as const;
export const TYPE_INDEX_LIMITS = {
  name: 256,
  documentation: 20_000,
  types: 50_000,
  members: 10_000,
  parameters: 1_024,
} as const;

const GUID_PATTERN = /^\{?[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$/;
const CONTROL_CHARACTERS = /[\u0000-\u001F\u007F]/;

type UnknownRecord = Record<string, unknown>;
type TypeOfKind = (typeof TYPE_INDEX_KIND_VALUES)[number];
type MemberOfKind = (typeof TYPE_INDEX_MEMBER_KIND_VALUES)[number];
type IndexType = PortableTypeIndex['types'][number];
type IndexMember = IndexType['members'][number];

function fail(path: string, message: string): never {
  throw new Error(path ? `${path}：${message}` : message);
}

function describe(value: unknown): string {
  if (value === null) { return 'null'; }
  if (Array.isArray(value)) { return '数组'; }
  if (typeof value === 'string') { return '字符串'; }
  if (typeof value === 'number') { return '数字'; }
  if (typeof value === 'boolean') { return '布尔值'; }
  return typeof value;
}

function isRecord(value: unknown): value is UnknownRecord {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) { return false; }
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function objectAt(value: unknown, path: string): UnknownRecord {
  if (!isRecord(value)) { fail(path, `必须是 JSON 对象，实际为 ${describe(value)}。`); }
  return value;
}

function rejectExtra(object: UnknownRecord, allowed: readonly string[], path: string): void {
  for (const key of Object.keys(object)) {
    if (!allowed.includes(key)) { fail(path ? `${path}.${key}` : key, '不支持的字段。'); }
  }
}

function stringAt(value: unknown, path: string, maxLength: number): string {
  if (typeof value !== 'string') { fail(path, `必须是字符串，实际为 ${describe(value)}。`); }
  if (value.length > maxLength) { fail(path, `长度超过 ${maxLength} 个字符。`); }
  return value;
}

function nameAt(value: unknown, path: string): string {
  const text = stringAt(value, path, TYPE_INDEX_LIMITS.name);
  if (!text.trim()) { fail(path, '不能为空。'); }
  if (text !== text.trim()) { fail(path, '不能以空白开头或结尾。'); }
  if (CONTROL_CHARACTERS.test(text)) { fail(path, '不能包含控制字符。'); }
  return text;
}

function documentationAt(value: unknown, path: string): string {
  return stringAt(value, path, TYPE_INDEX_LIMITS.documentation);
}

function integerAt(value: unknown, path: string, max: number): number {
  if (typeof value !== 'number' || !Number.isInteger(value)) { fail(path, '必须是整数。'); }
  if (value < 0 || value > max) { fail(path, `必须在 0–${max} 之间。`); }
  return value;
}

function booleanAt(value: unknown, path: string): boolean {
  if (typeof value !== 'boolean') { fail(path, `必须是布尔值，实际为 ${describe(value)}。`); }
  return value;
}

function arrayAt(value: unknown, path: string, maxItems: number): unknown[] {
  if (!Array.isArray(value)) { fail(path, `必须是数组，实际为 ${describe(value)}。`); }
  if (value.length > maxItems) { fail(path, `元素数量超过 ${maxItems}。`); }
  return value;
}

function optionalString(
  object: UnknownRecord,
  key: string,
  path: string,
  read: (value: unknown, path: string) => string,
): string | undefined {
  const value = object[key];
  return value === undefined ? undefined : read(value, path);
}

function kindAt<const T extends readonly string[]>(value: unknown, path: string, allowed: T): T[number] {
  if (typeof value !== 'string' || !(allowed as readonly string[]).includes(value)) {
    fail(path, `必须是以下值之一：${allowed.join('、')}。`);
  }
  return value as T[number];
}

function parameterAt(value: unknown, path: string): VbaParameter {
  const object = objectAt(value, path);
  rejectExtra(object, ['name', 'typeName', 'optional', 'byRef', 'paramArray'], path);
  const name = nameAt(object['name'], `${path}.name`);
  const typeName = nameAt(object['typeName'], `${path}.typeName`);
  const optional = booleanAt(object['optional'], `${path}.optional`);
  const byRef = booleanAt(object['byRef'], `${path}.byRef`);
  const parameter: VbaParameter = { name, typeName, optional, byRef };
  if (object['paramArray'] !== undefined) {
    // false 与缺省等价，序列化时省略，避免索引出现无意义的布尔字段。
    if (booleanAt(object['paramArray'], `${path}.paramArray`)) { parameter.paramArray = true; }
  }
  return parameter;
}

function memberAt(value: unknown, path: string): IndexMember {
  const object = objectAt(value, path);
  rejectExtra(object, ['name', 'kind', 'returnType', 'parameters', 'documentation'], path);
  const name = nameAt(object['name'], `${path}.name`);
  const kind: MemberOfKind = kindAt(object['kind'], `${path}.kind`, TYPE_INDEX_MEMBER_KIND_VALUES);
  const returnType = optionalString(object, 'returnType', `${path}.returnType`, nameAt);
  const documentation = optionalString(object, 'documentation', `${path}.documentation`, documentationAt);
  const member: IndexMember = { name, kind };
  if (returnType !== undefined) { member.returnType = returnType; }
  if (object['parameters'] !== undefined) {
    const parameters = arrayAt(object['parameters'], `${path}.parameters`, TYPE_INDEX_LIMITS.parameters)
      .map((item, index) => parameterAt(item, `${path}.parameters[${index}]`));
    if (parameters.length > 0) { member.parameters = parameters; }
  }
  if (documentation !== undefined) { member.documentation = documentation; }
  return member;
}

function typeAt(value: unknown, path: string): IndexType {
  const object = objectAt(value, path);
  rejectExtra(object, ['name', 'kind', 'documentation', 'members'], path);
  const name = nameAt(object['name'], `${path}.name`);
  const kind: TypeOfKind = kindAt(object['kind'], `${path}.kind`, TYPE_INDEX_KIND_VALUES);
  const documentation = optionalString(object, 'documentation', `${path}.documentation`, documentationAt);
  const members = arrayAt(object['members'], `${path}.members`, TYPE_INDEX_LIMITS.members)
    .map((item, index) => memberAt(item, `${path}.members[${index}]`));
  return documentation === undefined ? { name, kind, members } : { name, kind, documentation, members };
}

/**
 * 校验任意 JSON 值并返回规范化副本。未知字段、类型错误与超限值一律拒绝。
 * 返回的对象只含契约字段，调用方可安全地用于序列化与语义分析。
 */
export function validateTypeIndex(value: unknown): PortableTypeIndex {
  const root = objectAt(value, '');
  rejectExtra(root, ['formatVersion', 'library', 'types'], '');
  if (root['formatVersion'] !== TYPE_INDEX_FORMAT_VERSION) {
    fail('formatVersion', `必须为 ${TYPE_INDEX_FORMAT_VERSION}。`);
  }
  const libraryValue = objectAt(root['library'], 'library');
  rejectExtra(libraryValue, ['name', 'guid', 'major', 'minor', 'lcid', 'bitness'], 'library');
  const libraryName = nameAt(libraryValue['name'], 'library.name');
  const guid = stringAt(libraryValue['guid'], 'library.guid', 64);
  if (!GUID_PATTERN.test(guid)) {
    fail('library.guid', '必须是形如 {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} 的 GUID。');
  }
  const major = integerAt(libraryValue['major'], 'library.major', 65_535);
  const minor = integerAt(libraryValue['minor'], 'library.minor', 65_535);
  const lcid = integerAt(libraryValue['lcid'], 'library.lcid', 4_294_967_295);
  const bitness = integerAt(libraryValue['bitness'], 'library.bitness', 64);
  if (bitness !== 32 && bitness !== 64) { fail('library.bitness', '必须为 32 或 64。'); }
  const types = arrayAt(root['types'], 'types', TYPE_INDEX_LIMITS.types)
    .map((item, index) => typeAt(item, `types[${index}]`));
  return {
    formatVersion: TYPE_INDEX_FORMAT_VERSION,
    library: { name: libraryName, guid, major, minor, lcid, bitness },
    types,
  };
}

/** 解析 UTF-8 文本（允许 BOM）并校验。错误信息保留 JSON 解析原因。 */
export function parseTypeIndex(text: string): PortableTypeIndex {
  if (typeof text !== 'string') { throw new Error('类型索引内容必须是字符串。'); }
  const content = text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
  let value: unknown;
  try {
    value = JSON.parse(content);
  } catch (error) {
    throw new Error(`类型索引不是有效 JSON：${error instanceof Error ? error.message : '解析失败。'}`);
  }
  return validateTypeIndex(value);
}

/** 规范化为两空格缩进、LF 换行、以换行结尾的 UTF-8 文本。先校验再写出。 */
export function stringifyTypeIndex(index: PortableTypeIndex): string {
  return `${JSON.stringify(validateTypeIndex(index), null, 2)}\n`;
}

/**
 * 计算建议文件名。名称、版本、位数与 LCID 均参与命名，避免不同版本静默覆盖。
 * 例如 Excel 1.8 64 位 LCID 1033 → Excel-1.8-64bit-1033.json。
 */
export function typeIndexFileName(index: PortableTypeIndex): string {
  const base = index.library.name
    .replace(/[^\p{L}\p{N}._-]+/gu, '-')
    .replace(/^[-.]+|[-.]+$/g, '') || 'type-library';
  return `${base}-${index.library.major}.${index.library.minor}-${index.library.bitness}bit-${index.library.lcid}.json`;
}

/** 索引身份：忽略 GUID 大小写与花括号差异，保留名称、版本、LCID 与位数。 */
export function typeIndexIdentity(index: PortableTypeIndex): string {
  const guid = index.library.guid.replace(/[{}]/g, '').toLowerCase();
  return `${index.library.name.toLowerCase()}|${guid}|${index.library.major}.${index.library.minor}|${index.library.lcid}|${index.library.bitness}`;
}

/** 面向消息与日志的简短描述。 */
export function typeIndexLabel(index: PortableTypeIndex): string {
  return `${index.library.name} ${index.library.major}.${index.library.minor}（${index.library.bitness} 位，LCID ${index.library.lcid}）`;
}
