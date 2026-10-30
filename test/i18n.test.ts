import assert from 'node:assert/strict';
import test from 'node:test';
import {
  MESSAGE_KEYS,
  LOCALES,
  catalogFor,
  formatMessage,
  normalizeLanguageTag,
  translate,
} from '../src/i18n/catalog';

test('语言标签归一化：中文简繁与回退', () => {
  assert.equal(normalizeLanguageTag('zh-cn'), 'zh-cn');
  assert.equal(normalizeLanguageTag('zh-hans'), 'zh-cn');
  assert.equal(normalizeLanguageTag('zh'), 'zh-cn');
  assert.equal(normalizeLanguageTag('zh-tw'), 'zh-tw');
  assert.equal(normalizeLanguageTag('zh-hk'), 'zh-tw');
  assert.equal(normalizeLanguageTag('zh-hant'), 'zh-tw');
  assert.equal(normalizeLanguageTag('ja'), 'ja');
  assert.equal(normalizeLanguageTag('ja-JP'), 'ja');
  assert.equal(normalizeLanguageTag('ko-KR'), 'ko');
  assert.equal(normalizeLanguageTag('en-US'), 'en');
  assert.equal(normalizeLanguageTag('fr'), 'en');
  assert.equal(normalizeLanguageTag(''), 'en');
  assert.equal(normalizeLanguageTag(undefined), 'en');
});

test('所有语言都提供全部文案且非空', () => {
  assert.ok(MESSAGE_KEYS.length > 40);
  for (const locale of LOCALES) {
    const catalog = catalogFor(locale);
    for (const key of MESSAGE_KEYS) {
      const value = catalog[key];
      assert.equal(typeof value, 'string', `${locale} 缺少 ${key}`);
      assert.ok(value.length > 0, `${locale} 的 ${key} 为空`);
    }
  }
});

test('占位符在各语言之间保持一致', () => {
  const english = catalogFor('en');
  const placeholders = (text: string): string[] => [...text.matchAll(/\{(\d+)\}/g)].map(match => match[1]!).sort();
  for (const locale of LOCALES) {
    const catalog = catalogFor(locale);
    for (const key of MESSAGE_KEYS) {
      assert.deepEqual(placeholders(catalog[key]), placeholders(english[key]), `${locale} 的 ${key} 占位符不一致`);
    }
  }
});

test('文案格式化与回退', () => {
  assert.equal(formatMessage('a {0} b {1}', ['x', 2]), 'a x b 2');
  assert.equal(translate('zh-cn', 'go.test.success', 'model', 'text'), '连接成功。模型：model\ntext');
  assert.equal(translate('en', 'go.test.success', 'model', 'text'), 'Connection OK. Model: model\ntext');
});
