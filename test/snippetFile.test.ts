import assert from 'node:assert/strict';
import test from 'node:test';
import {
  CUSTOM_SNIPPET_TEMPLATE,
  SnippetFileError,
  matchSnippets,
  parseCustomSnippets,
  stripJsonComments,
} from '../src/snippetFile';

test('注释剥离保留字符串中的斜杠', () => {
  const text = '{\n  // comment\n  "a": "http://example.com/*x*/", /* block */\n  "b": 1\n}';
  const stripped = stripJsonComments(text);
  assert.ok(!stripped.includes('comment'));
  assert.ok(!stripped.includes('block'));
  assert.ok(stripped.includes('http://example.com/*x*/'));
  assert.deepEqual(JSON.parse(stripped), { a: 'http://example.com/*x*/', b: 1 });
});

test('解析自定义片段：字符串或数组前缀、正文按行连接', () => {
  const snippets = parseCustomSnippets(`{
    // 注释
    "If block": { "prefix": ["if", "ifthen"], "body": ["If \${1:condition} Then", "\\t\${2:}", "End If"], "description": "if" },
    "Single": { "prefix": "sgl", "body": "MsgBox 1" }
  }`);
  assert.equal(snippets.length, 2);
  const first = snippets[0]!;
  assert.deepEqual(first.prefixes, ['if', 'ifthen']);
  assert.equal(first.body, 'If ${1:condition} Then\n\t${2:}\nEnd If');
  assert.equal(snippets[1]!.body, 'MsgBox 1');
});

test('无效片段文件给出可读错误', () => {
  assert.throws(() => parseCustomSnippets('not json'), SnippetFileError);
  assert.throws(() => parseCustomSnippets('[]'), SnippetFileError);
  assert.throws(() => parseCustomSnippets('{"x": {"prefix": "a"}}'), SnippetFileError);
  assert.throws(() => parseCustomSnippets('{"x": {"body": ["a"]}}'), SnippetFileError);
});

test('前缀匹配不区分大小写且支持部分输入', () => {
  const snippets = parseCustomSnippets('{"a": {"prefix": ["If", "ifthen"], "body": ["x"]}, "b": {"prefix": "for", "body": ["y"]}}');
  assert.equal(matchSnippets(snippets, 'i').length, 1);
  assert.equal(matchSnippets(snippets, 'IF').length, 1);
  assert.equal(matchSnippets(snippets, 'for').length, 1);
  assert.equal(matchSnippets(snippets, 'z').length, 0);
  assert.equal(matchSnippets(snippets, '').length, 0);
});

test('新建片段模板可直接解析', () => {
  const snippets = parseCustomSnippets(CUSTOM_SNIPPET_TEMPLATE);
  assert.equal(snippets.length, 1);
  assert.equal(snippets[0]!.prefixes[0], 'log');
  assert.match(snippets[0]!.body, /\$\{1:message\}/);
});
