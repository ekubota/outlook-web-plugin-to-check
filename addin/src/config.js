/* global window */
// ビルド時に scripts/build.js が __BASE_URL__ を実際の Cloud Run URL に置換します。
window.DOMAIN_GUARD_CONFIG = {
  // API のベース URL（アドインの配信元と同一オリジンにしておくと CORS 不要）
  apiBaseUrl: '__BASE_URL__',

  // API 呼び出しのタイムアウト（ミリ秒）。5 秒を超えると Outlook が
  // 「予想以上に時間が掛かっています」ダイアログを出すため、それより短くしておく
  timeoutMs: 3500,

  // API 呼び出しに失敗したときの挙動
  //   'prompt' : 確認ダイアログ（このまま送信 / 送信しない）を出す（既定）
  //   'allow'  : そのまま送信させる
  failMode: 'prompt',

  // 確認ダイアログに個別表示する宛先の最大件数（本文は 500 文字まで）
  maxRecipientsInMessage: 8,

  // サーバ側で API キーを設定した場合のみ指定（クライアントに埋め込まれる点に注意）
  apiKey: '',
};
