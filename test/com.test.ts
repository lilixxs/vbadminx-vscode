import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import {
  TYPE_INDEX_FORMAT_VERSION,
  TYPE_INDEX_KIND_VALUES,
  TYPE_INDEX_MEMBER_KIND_VALUES,
  parseTypeIndex,
  stringifyTypeIndex,
  typeIndexFileName,
  typeIndexIdentity,
  typeIndexLabel,
  validateTypeIndex,
} from '../src/typeIndex.js';

/* eslint-disable @typescript-eslint/no-explicit-any -- 测试刻意构造无效 JSON 值。 */
function fixture(): any {
  return {
    formatVersion: 1,
    library: {
      name: 'Excel',
      guid: '{00020813-0000-0000-C000-000000000046}',
      major: 1,
      minor: 8,
      lcid: 0,
      bitness: 64,
    },
    types: [
      {
        name: 'Application',
        kind: 'class',
        documentation: 'Excel 应用程序对象。',
        members: [
          { name: 'Workbooks', kind: 'property', returnType: 'Workbooks', documentation: '打开的工作簿集合。' },
          {
            name: 'Calculate',
            kind: 'method',
            parameters: [
              { name: 'Mode', typeName: 'Variant', optional: true, byRef: false },
              { name: 'Rows', typeName: 'Variant', optional: false, byRef: true, paramArray: true },
            ],
          },
          { name: 'xlA1', kind: 'constant', returnType: 'Long' },
        ],
      },
      {
        name: 'XlReferenceStyle',
        kind: 'enum',
        members: [{ name: 'xlA1', kind: 'constant', returnType: 'Long' }],
      },
    ],
  };
}

function schemaFile(): string {
  const candidates = [
    path.resolve(__dirname, '../../schemas/type-index.schema.json'),
    path.resolve(process.cwd(), 'schemas/type-index.schema.json'),
  ];
  const found = candidates.find(candidate => existsSync(candidate));
  assert.ok(found, '找不到 schemas/type-index.schema.json');
  return found;
}

test('接受有效索引并返回不含共享引用的规范化副本', () => {
  const source = fixture();
  const index = validateTypeIndex(source);
  assert.equal(index.formatVersion, 1);
  assert.deepEqual(index, JSON.parse(JSON.stringify(source)));
  assert.notEqual(index, source);
  assert.notEqual(index.library, source.library);
  assert.notEqual(index.types, source.types);
  const sourceParameters = source.types[0].members[1].parameters;
  assert.notEqual(index.types[0]!.members[1]!.parameters, sourceParameters);
  assert.equal(index.types[0]!.members[1]!.parameters![1]!.paramArray, true);
});

test('拒绝错误的格式版本', () => {
  for (const value of [2, '1', null, undefined, true]) {
    const input = fixture();
    input.formatVersion = value;
    assert.throws(() => validateTypeIndex(input), /formatVersion：必须为 1/, String(value));
  }
});

test('拒绝根、library、type、member、parameter 上的未知字段', () => {
  const cases: Array<[string, (input: any) => void, RegExp]> = [
    ['根', input => { input.extra = 1; }, /extra：不支持的字段/],
    ['library', input => { input.library.extra = 1; }, /library\.extra：不支持的字段/],
    ['type', input => { input.types[0].extra = 1; }, /types\[0\]\.extra：不支持的字段/],
    ['member', input => { input.types[0].members[0].extra = 1; }, /types\[0\]\.members\[0\]\.extra：不支持的字段/],
    ['parameter', input => { input.types[0].members[1].parameters[0].extra = 1; }, /parameters\[0\]\.extra：不支持的字段/],
  ];
  for (const [label, mutate, pattern] of cases) {
    const input = fixture();
    mutate(input);
    assert.throws(() => validateTypeIndex(input), pattern, label);
  }
});

test('校验 library 的标识字段', () => {
  const cases: Array<[string, (input: any) => void, RegExp]> = [
    ['name 缺失', input => { delete input.library.name; }, /library\.name：必须是字符串，实际为 undefined/],
    ['name 为空', input => { input.library.name = ''; }, /library\.name：不能为空/],
    ['name 含首尾空白', input => { input.library.name = ' Excel'; }, /library\.name：不能以空白开头或结尾/],
    ['name 含控制字符', input => { input.library.name = 'Ex\ncel'; }, /library\.name：不能包含控制字符/],
    ['guid 格式错误', input => { input.library.guid = 'not-a-guid'; }, /library\.guid：必须是形如/],
    ['guid 参数类型错误', input => { input.library.guid = 1; }, /library\.guid：必须是字符串/],
    ['major 为负', input => { input.library.major = -1; }, /library\.major：必须在 0–65535/],
    ['major 为小数', input => { input.library.major = 1.5; }, /library\.major：必须是整数/],
    ['minor 超限', input => { input.library.minor = 65_536; }, /library\.minor：必须在 0–65535/],
    ['lcid 为负', input => { input.library.lcid = -1; }, /library\.lcid：必须在 0–4294967295/],
    ['lcid 超限', input => { input.library.lcid = 4_294_967_296; }, /library\.lcid：必须在 0–4294967295/],
    ['bitness 非 32/64', input => { input.library.bitness = 16; }, /library\.bitness：必须为 32 或 64/],
  ];
  for (const [label, mutate, pattern] of cases) {
    const input = fixture();
    mutate(input);
    assert.throws(() => validateTypeIndex(input), pattern, label);
  }

  for (const guid of ['{00020813-0000-0000-C000-000000000046}', '00020813-0000-0000-c000-000000000046']) {
    const input = fixture();
    input.library.guid = guid;
    assert.equal(validateTypeIndex(input).library.guid, guid);
  }
  for (const bitness of [32, 64]) {
    const input = fixture();
    input.library.bitness = bitness;
    assert.equal(validateTypeIndex(input).library.bitness, bitness);
  }
  const maxLcid = fixture();
  maxLcid.library.lcid = 4_294_967_295;
  assert.equal(validateTypeIndex(maxLcid).library.lcid, 4_294_967_295);
});

test('校验类型与成员的枚举值', () => {
  const cases: Array<[string, (input: any) => void, RegExp]> = [
    ['types 不是数组', input => { input.types = {}; }, /types：必须是数组/],
    ['类型名缺失', input => { delete input.types[0].name; }, /types\[0\]\.name：必须是字符串/],
    ['类型种类未知', input => { input.types[0].kind = 'struct'; }, /types\[0\]\.kind：必须是以下值之一/],
    ['成员种类未知', input => { input.types[0].members[0].kind = 'field'; }, /members\[0\]\.kind：必须是以下值之一/],
    ['成员名称为空', input => { input.types[0].members[0].name = ''; }, /members\[0\]\.name：不能为空/],
    ['returnType 为空', input => { input.types[0].members[0].returnType = ''; }, /returnType：不能为空/],
    ['documentation 为 null', input => { input.types[0].documentation = null; }, /documentation：必须是字符串，实际为 null/],
  ];
  for (const [label, mutate, pattern] of cases) {
    const input = fixture();
    mutate(input);
    assert.throws(() => validateTypeIndex(input), pattern, label);
  }

  for (const kind of TYPE_INDEX_KIND_VALUES) {
    const input = fixture();
    input.types[0].kind = kind;
    assert.equal(validateTypeIndex(input).types[0]!.kind, kind);
  }
  for (const kind of TYPE_INDEX_MEMBER_KIND_VALUES) {
    const input = fixture();
    input.types[0].members[0].kind = kind;
    assert.equal(validateTypeIndex(input).types[0]!.members[0]!.kind, kind);
  }
});

test('校验参数形状', () => {
  const cases: Array<[string, (input: any) => void, RegExp]> = [
    ['缺少 optional', input => { delete input.types[0].members[1].parameters[0].optional; }, /parameters\[0\]\.optional：必须是布尔值/],
    ['byRef 不是布尔值', input => { input.types[0].members[1].parameters[0].byRef = 'true'; }, /parameters\[0\]\.byRef：必须是布尔值/],
    ['paramArray 不是布尔值', input => { input.types[0].members[1].parameters[1].paramArray = 'yes'; }, /parameters\[1\]\.paramArray：必须是布尔值/],
    ['参数名称为空', input => { input.types[0].members[1].parameters[0].name = ' '; }, /parameters\[0\]\.name：不能为空/],
    ['类型名为空', input => { input.types[0].members[1].parameters[0].typeName = ''; }, /parameters\[0\]\.typeName：不能为空/],
  ];
  for (const [label, mutate, pattern] of cases) {
    const input = fixture();
    mutate(input);
    assert.throws(() => validateTypeIndex(input), pattern, label);
  }

  const input = fixture();
  input.types[0].members[1].parameters[1].paramArray = false;
  const index = validateTypeIndex(input);
  assert.equal(index.types[0]!.members[1]!.parameters![1]!.paramArray, undefined);

  const empty = fixture();
  empty.types[0].members[0].parameters = [];
  assert.equal(validateTypeIndex(empty).types[0]!.members[0]!.parameters, undefined);
});

test('执行数量与长度上限', () => {
  const longName = fixture();
  longName.library.name = 'x'.repeat(257);
  assert.throws(() => validateTypeIndex(longName), /library\.name：长度超过 256/);

  const maxName = fixture();
  maxName.library.name = 'x'.repeat(256);
  assert.equal(validateTypeIndex(maxName).library.name.length, 256);

  const longDocumentation = fixture();
  longDocumentation.types[0].documentation = 'x'.repeat(20_001);
  assert.throws(() => validateTypeIndex(longDocumentation), /documentation：长度超过 20000/);

  const maxDocumentation = fixture();
  maxDocumentation.types[0].documentation = 'x'.repeat(20_000);
  assert.equal(validateTypeIndex(maxDocumentation).types[0]!.documentation!.length, 20_000);

  const manyTypes = fixture();
  manyTypes.types = Array.from({ length: 50_001 }, (_, index) => ({ name: `T${index}`, kind: 'class', members: [] }));
  assert.throws(() => validateTypeIndex(manyTypes), /types：元素数量超过 50000/);

  const manyMembers = fixture();
  manyMembers.types[0].members = Array.from({ length: 10_001 }, (_, index) => ({ name: `M${index}`, kind: 'method' }));
  assert.throws(() => validateTypeIndex(manyMembers), /types\[0\]\.members：元素数量超过 10000/);

  const manyParameters = fixture();
  manyParameters.types[0].members[0].parameters = Array.from({ length: 1_025 }, (_, index) => ({
    name: `P${index}`, typeName: 'Variant', optional: false, byRef: false,
  }));
  assert.throws(() => validateTypeIndex(manyParameters), /parameters：元素数量超过 1024/);
});

test('parseTypeIndex 接受 BOM 并报告 JSON 错误', () => {
  const text = JSON.stringify(fixture());
  assert.deepEqual(parseTypeIndex(`\uFEFF${text}`), validateTypeIndex(fixture()));
  assert.throws(() => parseTypeIndex('{'), /类型索引不是有效 JSON/);
  assert.throws(() => parseTypeIndex('[]'), /必须是 JSON 对象，实际为 数组/);
  assert.throws(() => parseTypeIndex(JSON.stringify({ ...fixture(), formatVersion: 0 })), /formatVersion：必须为 1/);
});

test('stringifyTypeIndex 输出规范 UTF-8 文本并重新校验', () => {
  const index = validateTypeIndex(fixture());
  const text = stringifyTypeIndex(index);
  assert.ok(text.endsWith('\n'));
  assert.ok(!text.includes('\r'));
  assert.ok(text.includes('\n  "formatVersion": 1'));
  assert.notEqual(text.charCodeAt(0), 0xFEFF);
  assert.deepEqual(parseTypeIndex(text), index);
  assert.throws(() => stringifyTypeIndex({ ...fixture(), formatVersion: 2 }), /formatVersion：必须为 1/);
});

test('文件名包含名称、版本、位数与 LCID，且清理非法字符', () => {
  const index = validateTypeIndex(fixture());
  assert.equal(typeIndexFileName(index), 'Excel-1.8-64bit-0.json');
  const office = fixture();
  office.library.name = 'Office / Excel';
  assert.equal(typeIndexFileName(validateTypeIndex(office)), 'Office-Excel-1.8-64bit-0.json');
  const chinese = fixture();
  chinese.library.name = '测试 库';
  assert.equal(typeIndexFileName(validateTypeIndex(chinese)), '测试-库-1.8-64bit-0.json');
  const lcid = fixture();
  lcid.library.lcid = 1033;
  assert.equal(typeIndexFileName(validateTypeIndex(lcid)), 'Excel-1.8-64bit-1033.json');
});

test('索引身份忽略 GUID 大小写与花括号，但区分版本与位数', () => {
  const base = validateTypeIndex(fixture());
  const upper = validateTypeIndex(fixture());
  upper.library.guid = '{00020813-0000-0000-C000-000000000046}'.toUpperCase();
  const bare = validateTypeIndex(fixture());
  bare.library.guid = '00020813-0000-0000-c000-000000000046';
  assert.equal(typeIndexIdentity(base), typeIndexIdentity(upper));
  assert.equal(typeIndexIdentity(base), typeIndexIdentity(bare));
  const otherMinor = validateTypeIndex(fixture());
  otherMinor.library.minor = 9;
  const otherBitness = validateTypeIndex(fixture());
  otherBitness.library.bitness = 32;
  const otherLcid = validateTypeIndex(fixture());
  otherLcid.library.lcid = 1033;
  for (const other of [otherMinor, otherBitness, otherLcid]) {
    assert.notEqual(typeIndexIdentity(base), typeIndexIdentity(other));
  }
  assert.match(typeIndexLabel(base), /^Excel 1\.8（64 位，LCID 0）$/);
});

test('schema 与校验器保持一致', () => {
  const schema: any = JSON.parse(readFileSync(schemaFile(), 'utf8'));
  assert.equal(schema.$schema, 'https://json-schema.org/draft/2020-12/schema');
  assert.equal(schema.type, 'object');
  assert.equal(schema.additionalProperties, false);
  assert.deepEqual(schema.required, ['formatVersion', 'library', 'types']);
  assert.equal(schema.properties.formatVersion.const, TYPE_INDEX_FORMAT_VERSION);
  assert.deepEqual(schema.$defs.type.properties.kind.enum, [...TYPE_INDEX_KIND_VALUES]);
  assert.deepEqual(schema.$defs.member.properties.kind.enum, [...TYPE_INDEX_MEMBER_KIND_VALUES]);
  assert.deepEqual(schema.$defs.library.properties.bitness.enum, [32, 64]);
  assert.deepEqual(schema.$defs.library.required, ['name', 'guid', 'major', 'minor', 'lcid', 'bitness']);
  assert.deepEqual(schema.$defs.type.required, ['name', 'kind', 'members']);
  assert.deepEqual(schema.$defs.member.required, ['name', 'kind']);
  assert.deepEqual(schema.$defs.parameter.required, ['name', 'typeName', 'optional', 'byRef']);
  assert.equal(schema.properties.types.maxItems, 50_000);
  assert.equal(schema.$defs.type.properties.members.maxItems, 10_000);
  assert.equal(schema.$defs.member.properties.parameters.maxItems, 1_024);
  assert.equal(schema.$defs.name.maxLength, 256);
  assert.equal(schema.$defs.documentation.maxLength, 20_000);
  for (const definition of [schema, schema.$defs.library, schema.$defs.type, schema.$defs.member, schema.$defs.parameter]) {
    assert.equal(definition.additionalProperties, false);
  }

  const guidPattern = new RegExp(schema.$defs.guid.pattern);
  assert.ok(guidPattern.test('{00020813-0000-0000-C000-000000000046}'));
  assert.ok(guidPattern.test('00020813-0000-0000-c000-000000000046'));
  assert.ok(!guidPattern.test('{00020813-0000-0000-C000-00000000004}'));
  const namePattern = new RegExp(schema.$defs.name.pattern);
  assert.ok(namePattern.test('Workbooks'));
  assert.ok(namePattern.test('工作簿 集合'));
  assert.ok(!namePattern.test(' Workbooks'));
  assert.ok(!namePattern.test('Workbooks '));
  assert.ok(!namePattern.test(''));
});
