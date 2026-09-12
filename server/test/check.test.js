'use strict';

const test = require('node:test');
const assert = require('node:assert');

process.env.ALLOWED_DOMAINS = 'example.com,*.contoso.co.jp';
process.env.AUTO_ALLOW_SENDER_DOMAIN = 'true';

const { handleCheck } = require('../index');

function fakeReq(body, { method = 'POST', headers = {} } = {}) {
  return {
    method,
    path: '/api/check',
    body,
    headers,
    query: {},
    get: (name) => headers[String(name).toLowerCase()],
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

test('許可ドメインのみなら blocked は空', async () => {
  const res = fakeRes();
  await handleCheck(fakeReq({ recipients: [{ address: 'a@example.com', type: 'to' }] }), res);
  assert.equal(res.statusCode, 200);
  assert.equal(res.body.counts.blocked, 0);
  assert.equal(res.body.counts.allowed, 1);
});

test('許可リスト外のドメインを検出する', async () => {
  const res = fakeRes();
  await handleCheck(
    fakeReq({
      to: [{ address: 'a@example.com' }],
      cc: [{ address: 'b@gmail.com', displayName: '外部 太郎' }],
      bcc: ['c@mail.contoso.co.jp'],
    }),
    res
  );
  assert.deepEqual(res.body.blockedDomains, ['gmail.com']);
  assert.equal(res.body.blocked[0].type, 'cc');
  assert.equal(res.body.blocked[0].displayName, '外部 太郎');
  assert.equal(res.body.counts.allowed, 2, 'apex は許可、*.contoso.co.jp のサブドメインも許可');
});

test('アドレス未解決の宛先はブロック扱い', async () => {
  const res = fakeRes();
  await handleCheck(fakeReq({ recipients: [{ address: '/o=ExchangeLabs/cn=x', type: 'to' }] }), res);
  assert.equal(res.body.blocked[0].reason, 'unresolved_address');
});

test('送信者のドメインは自動的に許可される', async () => {
  const res = fakeRes();
  await handleCheck(
    fakeReq({ sender: 'me@mycorp.jp', recipients: [{ address: 'peer@mycorp.jp', type: 'to' }] }),
    res
  );
  assert.equal(res.body.counts.blocked, 0);
});

test('同一宛先の重複は 1 件にまとめる', async () => {
  const res = fakeRes();
  await handleCheck(
    fakeReq({ to: [{ address: 'x@gmail.com' }, { address: 'X@Gmail.com' }] }),
    res
  );
  assert.equal(res.body.blocked.length, 1);
});

test('GET は 405', async () => {
  const res = fakeRes();
  await handleCheck(fakeReq({}, { method: 'GET' }), res);
  assert.equal(res.statusCode, 405);
});
