/* global window */
// ビルド時に scripts/build.js が __BASE_URL__ を実際の Cloud Run URL に置換します。
window.DOMAIN_GUARD_CONFIG = {
  // API のベース URL（アドインの配信元と同一オリジンにしておくと CORS 不要）
  apiBaseUrl: '__BASE_URL__',

  // API 呼び出しのタイムアウト（ミリ秒）
  timeoutMs: 8000,

  // API 呼び出しに失敗したときの挙動
  //   'prompt' : Outlook 標準の「このまま送信 / 送信しない」ダイアログを出す（既定）
  //   'block'  : 送信を止める
  //   'allow'  : そのまま送信させる
  failMode: 'prompt',

  // ダイアログに個別表示する宛先の最大件数（超過分は「ほか N 件」と表示）
  maxRecipientsInDialog: 20,

  // サーバ側で API キーを設定した場合のみ指定（クライアントに埋め込まれる点に注意）
  apiKey: '',
};
