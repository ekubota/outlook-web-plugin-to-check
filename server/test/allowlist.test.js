'use strict';

const test = require('node:test');
const assert = require('node:assert');

const { domainOf, isAllowedDomain, matchesEntry, normalizeDomain, parseEntries, createAllowlistLoader } =
  require('../lib/allowlist');

test('normalizeDomain', () => {
  assert.equal(normalizeDomain('  Example.COM '), 'example.com');
  assert.equal(normalizeDomain('@Example.com'), 'example.com');
  assert.equal(normalizeDomain('example.com.'), 'example.com');
  assert.equal(normalizeDomain('user@sub.Example.com'), 'sub.example.com');
  assert.equal(normalizeDomain('not a domain'), '');
  assert.equal(normalizeDomain(null), '');
});

test('domainOf', () => {
  assert.equal(domainOf('taro@Example.co.jp'), 'example.co.jp');
  assert.equal(domainOf('"weird@name"@example.com'), 'example.com');
  assert.equal(domainOf('/o=ExchangeLabs/ou=...'), '', 'X500 アドレスはドメイン判定不可');
  assert.equal(domainOf(''), '');
});

test('matchesEntry: 完全一致とワイルドカード', () => {
  assert.ok(matchesEntry('example.com', 'example.com'));
  assert.ok(!matchesEntry('evil-example.com', 'example.com'));
  assert.ok(matchesEntry('mail.example.com', '*.example.com'));
  assert.ok(matchesEntry('a.b.example.com', '*.example.com'));
  assert.ok(!matchesEntry('example.com', '*.example.com'), 'apex は *. に一致しない');
  assert.ok(!matchesEntry('notexample.com', '*.example.com'));
});

test('isAllowedDomain', () => {
  const list = ['example.com', '*.contoso.co.jp'];
  assert.ok(isAllowedDomain('example.com', list));
  assert.ok(isAllowedDomain('mail.contoso.co.jp', list));
  assert.ok(!isAllowedDomain('contoso.co.jp', list));
  assert.ok(!isAllowedDomain('gmail.com', list));
  assert.ok(!isAllowedDomain('', list));
});

test('parseEntries: 配列・オブジェクト・CSV 文字列', () => {
  assert.deepEqual(parseEntries(['B.com', 'a.com', 'a.com']), ['a.com', 'b.com']);
  assert.deepEqual(parseEntries({ allowedDomains: ['X.com'] }), ['x.com']);
  assert.deepEqual(parseEntries('a.com, b.com\nc.com'), ['a.com', 'b.com', 'c.com']);
});

test('createAllowlistLoader: ALLOWED_DOMAINS が最優先', async () => {
  const loader = createAllowlistLoader({ ALLOWED_DOMAINS: 'foo.com, *.bar.com' });
  const list = await loader.get();
  assert.deepEqual(list.domains, ['*.bar.com', 'foo.com']);
  assert.equal(list.source, 'env:ALLOWED_DOMAINS');
});

test('createAllowlistLoader: 既定は同梱 JSON', async () => {
  const loader = createAllowlistLoader({});
  const list = await loader.get();
  assert.ok(list.domains.includes('example.com'));
  assert.match(list.source, /^file:/);
});
