/* global Office, window, fetch, setTimeout, clearTimeout, AbortController, console */

/**
 * OnMessageSend ハンドラー（Smart Alerts / イベントベースアクティブ化）。
 *
 *   1. 宛先（宛先・CC・BCC）を取得
 *   2. Cloud Run の /api/check に渡して、許可ドメイン一覧に無い宛先を判定
 *   3. 該当があれば独自ダイアログを出し、「送信する / 送信しない」を確認
 *   4. 結果を event.completed({ allowEvent }) で Outlook に返す
 */

// 既定値。__BASE_URL__ はビルド時に置換される（config.js を読み込めない
// classic Outlook の JavaScript ランタイムでも動くよう、ここにも持たせている）。
var DEFAULTS = {
  apiBaseUrl: '__BASE_URL__',
  timeoutMs: 8000,
  failMode: 'prompt',
  maxRecipientsInDialog: 20,
  apiKey: '',
};

var CONFIG = (typeof window !== 'undefined' && window.DOMAIN_GUARD_CONFIG) || {};

function config(key, fallback) {
  var value = CONFIG[key];
  if (value === undefined || value === '') value = DEFAULTS[key];
  return value === undefined || value === '' ? fallback : value;
}

function apiUrl(pathname) {
  var base = String(config('apiBaseUrl', '')).replace(/\/+$/, '');
  return base + pathname;
}

/** Office の コールバック API を Promise 化する。 */
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
  var timeoutMs = config('timeoutMs', 8000);
  var timer = setTimeout(function () {
    if (controller) controller.abort();
  }, timeoutMs);

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

/** ブロック対象をダイアログ URL のクエリに詰める。 */
function buildDialogUrl(result) {
  var limit = config('maxRecipientsInDialog', 20);
  var blocked = result.blocked || [];
  var payload = {
    domains: result.blockedDomains || [],
    recipients: blocked.slice(0, limit).map(function (b) {
      return { a: b.address, n: b.displayName, t: b.type, d: b.domain, r: b.reason };
    }),
    more: Math.max(0, blocked.length - limit),
    total: (result.counts && result.counts.total) || blocked.length,
  };
  return apiUrl('/src/dialog/dialog.html') + '?p=' + encodeURIComponent(JSON.stringify(payload));
}

/**
 * 独自ダイアログを表示し、'send' / 'cancel' を返す。
 * ダイアログを開けなかった場合は reject する（呼び出し側でフォールバック）。
 */
function askUser(result) {
  return new Promise(function (resolve, reject) {
    Office.context.ui.displayDialogAsync(
      buildDialogUrl(result),
      { height: 55, width: 42, displayInIframe: true },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error((asyncResult.error && asyncResult.error.message) || 'dialog failed'));
          return;
        }
        var dialog = asyncResult.value;
        var settled = false;
        var finish = function (answer) {
          if (settled) return;
          settled = true;
          try {
            dialog.close();
          } catch (e) {
            /* すでに閉じている */
          }
          resolve(answer);
        };

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var message = {};
          try {
            message = JSON.parse(arg.message);
          } catch (e) {
            /* 想定外のメッセージは cancel 扱い */
          }
          finish(message.action === 'send' ? 'send' : 'cancel');
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          // 12006: ユーザーがダイアログを閉じた → 送信しない
          finish('cancel');
        });
      }
    );
  });
}

function completeAllow(event) {
  event.completed({ allowEvent: true });
}

function completeBlock(event, message) {
  event.completed({
    allowEvent: false,
    errorMessage: String(message).slice(0, 1000),
  });
}

/** Outlook 標準の「このまま送信 / 送信しない」ダイアログにフォールバックする。 */
function completePrompt(event, message) {
  try {
    event.completed({
      allowEvent: false,
      errorMessage: String(message).slice(0, 1000),
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
  } catch (e) {
    // sendModeOverride 非対応クライアントでは SoftBlock のまま停止する。
    completeBlock(event, message);
  }
}

function handleFailure(event, err) {
  console.error('[domain-guard] check failed:', err && err.message);
  var message =
    '宛先ドメインの確認に失敗しました（' +
    ((err && err.message) || 'unknown error') +
    '）。内容を確認してから送信してください。';
  var mode = config('failMode', 'prompt');
  if (mode === 'allow') completeAllow(event);
  else if (mode === 'block') completeBlock(event, message);
  else completePrompt(event, message);
}

function summarize(result) {
  var domains = result.blockedDomains || [];
  var head = domains.slice(0, 5).join('、');
  if (domains.length > 5) head += ' ほか' + (domains.length - 5) + '件';
  return '許可リストにないドメイン宛の宛先が含まれています: ' + head;
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  getRecipients(item)
    .then(function (recipients) {
      if (!recipients.length) {
        completeAllow(event);
        return null;
      }
      return checkWithServer(recipients).then(function (result) {
        if (!result || !result.blocked || result.blocked.length === 0) {
          completeAllow(event);
          return null;
        }
        return askUser(result).then(
          function (answer) {
            if (answer === 'send') completeAllow(event);
            else completeBlock(event, '送信を中止しました。' + summarize(result));
          },
          function (dialogErr) {
            // ダイアログを開けなかったときは Outlook 標準の確認ダイアログに委ねる。
            console.error('[domain-guard] dialog failed:', dialogErr && dialogErr.message);
            completePrompt(event, summarize(result) + ' 送信してよいか確認してください。');
          }
        );
      });
    })
    .catch(function (err) {
      handleFailure(event, err);
    });
}

// Office ランタイムへ登録（マニフェストの FunctionName と一致させること）
if (typeof Office !== 'undefined' && Office.actions && Office.actions.associate) {
  Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
}
