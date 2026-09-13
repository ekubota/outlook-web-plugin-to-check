/* global Office, window, fetch, setTimeout, clearTimeout, AbortController, console */

/**
 * OnMessageSend ハンドラー（Smart Alerts / イベントベースアクティブ化）。
 *
 *   1. 宛先（宛先・CC・BCC）を取得
 *   2. Cloud Run の /api/check に渡して、許可ドメイン一覧に無い宛先を判定
 *   3. 該当があれば event.completed({ allowEvent: false, errorMessage }) を返す
 *      → マニフェストの SendMode="PromptUser" により、Outlook が
 *        「このまま送信 / 送信しない」の 2 ボタン付きダイアログを表示する
 *
 * 注意: イベントハンドラー内では displayDialogAsync 等の UI API は使用できない
 * （Microsoft の制限）ため、独自ダイアログではなく Smart Alerts 標準ダイアログを使う。
 */

// 既定値。__BASE_URL__ はビルド時に置換される（config.js を読み込めない
// classic Outlook の JavaScript ランタイムでも動くよう、ここにも持たせている）。
var DEFAULTS = {
  apiBaseUrl: '__BASE_URL__',
  timeoutMs: 3500,
  failMode: 'prompt',
  maxRecipientsInMessage: 8,
  apiKey: '',
  // 診断用: true のとき処理の各段階で /health?stage=... を呼ぶ（Cloud Run のログで追跡できる）
  debugBeacon: true,
};

var CONFIG = (typeof window !== 'undefined' && window.DOMAIN_GUARD_CONFIG) || {};

// Smart Alerts ダイアログの本文は 500 文字まで
var MAX_MESSAGE_LENGTH = 500;

// Outlook は 5 秒を超えると「予想以上に時間が掛かっています」を出すため、その前に必ず完了させる
var HARD_DEADLINE_MS = 4500;

function config(key, fallback) {
  var value = CONFIG[key];
  if (value === undefined || value === '') value = DEFAULTS[key];
  return value === undefined || value === '' ? fallback : value;
}

function apiUrl(pathname) {
  var base = String(config('apiBaseUrl', '')).replace(/\/+$/, '');
  return base + pathname;
}

/** 診断用ビーコン。失敗しても処理には影響させない。 */
function beacon(stage) {
  if (!config('debugBeacon', false)) return;
  try {
    var url = apiUrl('/health?stage=' + encodeURIComponent(String(stage).slice(0, 200)) + '&t=' + Date.now());
    fetch(url, { method: 'GET', keepalive: true, cache: 'no-store' }).catch(function () {});
  } catch (e) {
    /* ignore */
  }
}

/** Office のコールバック API を Promise 化する。 */
function asPromise(fn) {
  return new Promise(function (resolve, reject) {
    fn(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error((result.error && result.error.message) || 'Office API error'));
    });
  });
}

function getRecipients(item) {
  return Promise.all([
    asPromise(function (cb) {
      item.to.getAsync(cb);
    }),
    asPromise(function (cb) {
      item.cc.getAsync(cb);
    }),
    asPromise(function (cb) {
      item.bcc.getAsync(cb);
    }),
  ]).then(function (values) {
    var types = ['to', 'cc', 'bcc'];
    var out = [];
    values.forEach(function (list, i) {
      (list || []).forEach(function (r) {
        out.push({
          address: r.emailAddress || '',
          displayName: r.displayName || '',
          type: types[i],
        });
      });
    });
    return out;
  });
}

function senderAddress() {
  try {
    var profile = Office.context.mailbox.userProfile;
    return (profile && profile.emailAddress) || '';
  } catch (e) {
    return '';
  }
}

function checkWithServer(recipients) {
  var controller = typeof AbortController === 'function' ? new AbortController() : null;
  var timer = setTimeout(function () {
    if (controller) controller.abort();
  }, config('timeoutMs', 3500));

  var headers = { 'Content-Type': 'application/json' };
  var apiKey = config('apiKey', '');
  if (apiKey) headers['X-Api-Key'] = apiKey;

  return fetch(apiUrl('/api/check'), {
    method: 'POST',
    headers: headers,
    body: JSON.stringify({ recipients: recipients, sender: senderAddress() }),
    signal: controller ? controller.signal : undefined,
  })
    .then(function (res) {
      if (!res.ok) throw new Error('API responded ' + res.status);
      return res.json();
    })
    .then(function (json) {
      clearTimeout(timer);
      return json;
    })
    .catch(function (err) {
      clearTimeout(timer);
      throw err;
    });
}

var TYPE_LABEL = { to: '宛先', cc: 'CC', bcc: 'BCC' };

/** ブロック対象の宛先を、ダイアログ本文（500 文字以内）に整形する。 */
function buildMessage(result) {
  var blocked = result.blocked || [];
  var limit = config('maxRecipientsInMessage', 8);
  var lines = blocked.slice(0, limit).map(function (b) {
    var label = TYPE_LABEL[b.type] || b.type || '';
    var who = b.address || '(アドレス未解決)';
    return '・' + who + (label ? '（' + label + '）' : '');
  });
  if (blocked.length > limit) lines.push('・ほか ' + (blocked.length - limit) + ' 件');

  var head = '許可リストに登録されていないドメイン宛の宛先が含まれています。内容を確認してから送信してください。\n\n';
  var message = head + lines.join('\n');

  if (message.length > MAX_MESSAGE_LENGTH) {
    // 宛先を減らしてドメインの一覧だけにする
    var domains = (result.blockedDomains || []).slice(0, 10).join('、');
    message = head + '対象ドメイン: ' + domains;
    if ((result.blockedDomains || []).length > 10) message += ' ほか';
  }
  return message.slice(0, MAX_MESSAGE_LENGTH);
}

/**
 * event.completed を一度だけ呼ぶためのラッパー。
 * 期限タイマーと通常処理のどちらが先に来ても二重に完了させない。
 */
function createCompleter(event) {
  var done = false;
  return function complete(options, stage) {
    if (done) return;
    done = true;
    beacon('complete:' + stage);
    try {
      event.completed(options);
    } catch (e) {
      beacon('complete-threw:' + (e && e.message));
    }
  };
}

function allowOptions() {
  return { allowEvent: true };
}

/** SendMode="PromptUser" では、Outlook が「このまま送信 / 送信しない」を表示する。 */
function promptOptions(message) {
  return {
    allowEvent: false,
    errorMessage: String(message).slice(0, MAX_MESSAGE_LENGTH),
  };
}

function failureOptions(err) {
  if (config('failMode', 'prompt') === 'allow') return allowOptions();
  return promptOptions(
    '宛先ドメインの確認ができませんでした（' +
      ((err && err.message) || 'unknown error') +
      '）。宛先を確認してから送信してください。'
  );
}

function onMessageSendHandler(event) {
  beacon('handler-start');
  var complete = createCompleter(event);

  // 保険: どこかで止まっても 5 秒の閾値より前に必ず完了させる
  var deadline = setTimeout(function () {
    complete(failureOptions(new Error('timeout')), 'deadline');
  }, HARD_DEADLINE_MS);

  function finish(options, stage) {
    clearTimeout(deadline);
    complete(options, stage);
  }

  try {
    var item = Office.context.mailbox.item;
    beacon('item:' + (item ? 'ok' : 'null'));

    getRecipients(item)
      .then(function (recipients) {
        beacon('recipients:' + recipients.length);
        if (!recipients.length) {
          finish(allowOptions(), 'no-recipients');
          return null;
        }
        return checkWithServer(recipients).then(function (result) {
          var blocked = (result && result.blocked && result.blocked.length) || 0;
          beacon('checked:blocked=' + blocked);
          if (blocked === 0) finish(allowOptions(), 'allowed');
          else finish(promptOptions(buildMessage(result)), 'prompt');
        });
      })
      .catch(function (err) {
        console.error('[domain-guard] check failed:', err && err.message);
        beacon('error:' + ((err && err.message) || 'unknown'));
        finish(failureOptions(err), 'error');
      });
  } catch (err) {
    // 同期例外でも必ず完了させる（完了しないと Outlook が待ち続ける）
    beacon('sync-error:' + ((err && err.message) || 'unknown'));
    finish(failureOptions(err), 'sync-error');
  }
}

// ---- ランタイムへの登録 ---------------------------------------------------------

if (typeof window !== 'undefined') {
  window.addEventListener('error', function (e) {
    beacon('window-error:' + (e && e.message));
  });
  window.addEventListener('unhandledrejection', function (e) {
    beacon('unhandled-rejection:' + (e && e.reason && e.reason.message));
  });
}

beacon('script-loaded:office=' + (typeof Office !== 'undefined') + ',actions=' + !!(typeof Office !== 'undefined' && Office.actions));

// マニフェストの FunctionName と一致させること
if (typeof Office !== 'undefined' && Office.actions && Office.actions.associate) {
  Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
  beacon('associated');
} else {
  beacon('associate-unavailable');
}

// ホスト情報（どのクライアント・アカウント種別で動いているかの確認用）
if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
  Office.onReady(function (info) {
    var diag = 'ready:host=' + (info && info.host) + ',platform=' + (info && info.platform);
    try {
      var mb = Office.context.mailbox;
      diag += ',type=' + (mb && mb.userProfile && mb.userProfile.accountType);
      diag += ',req1.12=' + Office.context.requirements.isSetSupported('Mailbox', '1.12');
    } catch (e) {
      diag += ',diag-error=' + (e && e.message);
    }
    beacon(diag);
  });
}
