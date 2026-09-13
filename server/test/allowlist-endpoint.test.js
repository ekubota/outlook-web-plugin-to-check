'use strict';

const test = require('node:test');
const assert = require('node:assert');

process.env.ALLOWED_DOMAINS = 'example.com,*.contoso.co.jp';
delete process.env.ADMIN_API_KEY;

const { handleAllowlist } = require('../index');

function fakeReq(headers = {}, query = {}) {
  const lower = Object.fromEntries(Object.entries(headers).map(([k, v]) => [k.toLowerCase(), v]));
  return {
    method: 'GET',
    path: '/api/allowlist',
    headers: lower,
    query,
    get: (name) => lower[String(name).toLowerCase()],
  };
}

function fakeRes() {
  const res = {
    statusCode: 0,
    body: null,
    headers: {},
    set(k, v) {
      res.headers[k] = v;
      return res;
    },
    status(code) {
      res.statusCode = code;
      return res;
    },
    json(payload) {
      res.body = payload;
      return res;
    },
    send(payload) {
      res.body = payload;
      return res;
    },
  };
  return res;
}

async function call(headers) {
  const res = fakeRes();
  await handleAllowlist(fakeReq(headers), res);
  return res;
}

test('ADMIN_API_KEY 未設定なら、キーを付けても 404（エンドポイントが存在しない扱い）', async (t) => {
  delete process.env.ADMIN_API_KEY;
  assert.equal((await call({})).statusCode, 404);
  assert.equal((await call({ 'X-Admin-Key': 'anything' })).statusCode, 404);
  assert.equal((await call({ 'X-Admin-Key': '' })).statusCode, 404);
});

test('ADMIN_API_KEY 設定時: キーなし・不一致は 404、中身は返さない', async (t) => {
  process.env.ADMIN_API_KEY = 'correct-horse-battery-staple';
  t.after(() => delete process.env.ADMIN_API_KEY);

  const none = await call({});
  assert.equal(none.statusCode, 404);
  assert.equal(typeof none.body, 'string', 'JSON の許可リストを返していない');

  const wrong = await call({ 'X-Admin-Key': 'wrong' });
  assert.equal(wrong.statusCode, 404);

  const prefix = await call({ 'X-Admin-Key': 'correct-horse' });
  assert.equal(prefix.statusCode, 404, '前方一致では通らない');
});

test('ADMIN_API_KEY 設定時: 一致すれば許可リストを返す（キャッシュさせない）', async (t) => {
  process.env.ADMIN_API_KEY = 'correct-horse-battery-staple';
  t.after(() => delete process.env.ADMIN_API_KEY);

  const ok = await call({ 'X-Admin-Key': 'correct-horse-battery-staple' });
  assert.equal(ok.statusCode, 200);
  assert.deepEqual(ok.body.domains, ['*.contoso.co.jp', 'example.com']);
  assert.equal(ok.headers['Cache-Control'], 'no-store');
});

test('クライアント用の X-Api-Key では管理用エンドポイントに入れない', async (t) => {
  process.env.ADMIN_API_KEY = 'correct-horse-battery-staple';
  t.after(() => delete process.env.ADMIN_API_KEY);

  const res = await call({ 'X-Api-Key': 'correct-horse-battery-staple' });
  assert.equal(res.statusCode, 404);
});
