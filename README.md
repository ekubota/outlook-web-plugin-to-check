# outlook-web-plugin-to-check

Outlook（ブラウザー版 / 新しい Outlook / デスクトップ版）でメールを送信するときに、
**宛先ドメインを許可リストと照合し、リストにないドメインが含まれていれば確認ダイアログを出して
「このまま送信する / 送信しない」を選ばせる** アドインと、その判定を行う
**Google Cloud Run（Cloud Run functions）** のサーバー実装です。

```
送信ボタン
   ↓ OnMessageSend イベント（Smart Alerts）
アドイン: 宛先/CC/BCC を取得
   ↓ POST /api/check
Cloud Run: 許可リストと照合 → 許可外の宛先を返す
   ↓ 該当あり
アドイン: 独自ダイアログを表示
   ├─「このまま送信する」→ event.completed({ allowEvent: true })  → 送信される
   └─「送信しない」      → event.completed({ allowEvent: false }) → 送信が止まる
```

## 構成

| パス | 内容 |
| --- | --- |
| `addin/manifest.xml` | アドインのマニフェスト（`__BASE_URL__` はビルド時に置換） |
| `addin/src/launchevent/` | 送信時イベントハンドラー（判定 → ダイアログ → 可否決定） |
| `addin/src/dialog/` | 確認ダイアログの UI |
| `addin/src/config.js` | クライアント側設定（API URL、タイムアウト、失敗時の挙動） |
| `server/index.js` | Cloud Run functions 本体（`/api/check`、`/api/allowlist` と静的配信） |
| `server/lib/allowlist.js` | ドメイン正規化と許可リスト照合ロジック |
| `server/config/allowlist.json` | 同梱の許可ドメイン一覧 |
| `scripts/build.js` | `addin/` を `server/public/` にコピーし URL を埋め込む |
| `scripts/deploy.js` | ビルド → `gcloud run deploy` → URL 確定後に再デプロイ |

アドインの静的ファイルは同じ Cloud Run サービスから配信します。アドインと API が
**同一オリジン**になるため CORS 設定が不要で、管理する URL も 1 つで済みます。

## 前提条件

- Microsoft 365 の職場・学校アカウント（Exchange Online）
  - **Outlook.com の個人アカウント（Hotmail/live.com 等）は非対応です。**
    Smart Alerts（`OnMessageSend`）のサポートクライアント表はバックエンドが
    Exchange Online / Exchange Server のマイルボックスに限定されており、
    個人アカウントのバックエンドは対象外です
    （[公式ドキュメントのサポート表](https://learn.microsoft.com/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#supported-clients-and-platforms)）。
- Outlook 要件セット **Mailbox 1.12** 以上
  - Outlook on the web / 新しい Outlook for Windows: 対応済み
  - classic Outlook for Windows: 2206 (build 15330.20196) 以上
  - Outlook for Mac: 16.65 (22082700) 以上
  - **Outlook モバイル（Android/iOS）は送信時イベント非対応**
- Google Cloud プロジェクトと `gcloud` CLI
- Node.js 20 以上

## 1. デプロイ

```bash
# 依存関係
npm --prefix server install

# プロジェクトを指定してデプロイ（ビルド → deploy → URL 確定後に再デプロイ）
node scripts/deploy.js --project <PROJECT_ID> --region asia-northeast1
```

`scripts/deploy.js` は内部で次を実行します。

```bash
gcloud run deploy domain-guard \
  --source ./server \
  --function domainGuard \
  --base-image nodejs22 \
  --region asia-northeast1 \
  --allow-unauthenticated
```

完了すると `https://domain-guard-xxxxxxxx.asia-northeast1.run.app` のような URL が表示され、
その URL を埋め込んだ **`dist/manifest.xml`** が生成されます。

> 手動で行う場合は `node scripts/build.js <サービス URL>` → `gcloud run deploy ...` の順に実行してください。
> URL を事前に確定させたい場合は Cloud Run のカスタムドメインをマッピングし、
> `BASE_URL=https://mail-guard.example.com node scripts/deploy.js ...` のように指定します。

### 動作確認

```bash
curl https://<サービス URL>/health
curl https://<サービス URL>/api/allowlist
curl -X POST https://<サービス URL>/api/check \
  -H 'Content-Type: application/json' \
  -d '{"sender":"me@example.com","to":[{"address":"a@example.com"},{"address":"b@gmail.com"}]}'
```

## 2. Outlook への登録

サイドロード（自分のアカウントでの動作確認用）と、管理者配置（組織展開・強制用）の
2通りがあります。**`OnMessageSend` はサイドロードでも実際に発火します**
（Microsoft 公式の [Smart Alerts walkthrough](https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough#try-it-out)
もサイドロードして動作確認する手順です）。ただし、サイドロードしたアドインはユーザー自身が
無効化・削除できてしまうため、**組織として送信時チェックを強制したい場合は管理者配置が必須**です。

| 配置方法 | 自分のアカウントで発火するか | 他ユーザーに強制できるか |
| --- | --- | --- |
| サイドロード（ファイルから追加） | する | しない（各自無効化・削除できる） |
| Microsoft 365 管理センターで管理者配置 | する | する（組織全体に強制配布） |

### 動作確認する（サイドロード）

1. Outlook on the web を開く（Microsoft 365 の職場・学校アカウントでサインイン）
2. 「設定（歯車）」→「アドインを管理」（または新しい Outlook の「アドインを取得」）
3. 「マイ アドイン」→「カスタム アドイン」→「ファイルから追加」
4. `dist/manifest.xml` を選択
5. **ブラウザーをリロード**（イベントハンドラーの登録はリロード後に有効になります）

### 組織全体に配布・強制する

1. [Microsoft 365 管理センター](https://admin.microsoft.com/) →「設定」→「統合アプリ」→「カスタム アプリのアップロード」
2. 「App type」で **Office Add-in**（Unified manifest ではなく add-in 専用マニフェスト）を選択
3. `dist/manifest.xml` を選択してアップロード
4. 割り当て先（全員 / 特定のユーザー・グループ）を選択
5. 反映には最大 24 時間かかることがあります（通常は数時間〜）。対象ユーザーは Outlook をリロードしてください

## 3. 許可ドメインの運用

優先順位は次のとおりです（上にあるものが優先）。

| 方法 | 設定 | 更新方法 |
| --- | --- | --- |
| 環境変数 | `ALLOWED_DOMAINS="example.com,*.contoso.co.jp"` | `gcloud run services update` で即時反映 |
| Cloud Storage | `ALLOWLIST_GCS_URI="gs://my-bucket/allowlist.json"` | **再デプロイ不要**。JSON を差し替えるだけ（最大 5 分キャッシュ） |
| 同梱 JSON | `server/config/allowlist.json` | 再デプロイが必要 |

書式:

```json
{
  "version": "2026-09-12",
  "allowedDomains": ["example.com", "*.example.com", "contoso.co.jp"]
}
```

- `example.com` … 完全一致のみ
- `*.example.com` … サブドメインのみに一致（`mail.example.com` ○ / `example.com` ✕）
  - apex も許可したい場合は両方書いてください
- 判定はすべて小文字に正規化して行います
- `/o=ExchangeLabs/...` のような未解決アドレス（社内の配布リストなど）はドメインを判定できないため
  **ブロック対象（`reason: "unresolved_address"`）** として扱われます

Cloud Storage を使う場合は、Cloud Run のサービスアカウントに
`roles/storage.objectViewer` を付与してください。

## 設定項目

### サーバー（環境変数）

| 変数 | 既定値 | 説明 |
| --- | --- | --- |
| `ALLOWED_DOMAINS` | （なし） | カンマ区切りの許可ドメイン。最優先 |
| `ALLOWLIST_GCS_URI` | （なし） | `gs://bucket/allowlist.json` |
| `ALLOWLIST_CACHE_TTL_MS` | `300000` | 許可リストのキャッシュ時間 |
| `AUTO_ALLOW_SENDER_DOMAIN` | `true` | 送信者自身のドメインを自動的に許可 |
| `API_KEY` | （なし） | 設定すると `X-Api-Key` ヘッダーを要求 |
| `ALLOWED_ORIGINS` | （なし） | CORS 許可オリジン（同一オリジン配信なら不要） |
| `MAX_RECIPIENTS` | `500` | 1 リクエストで判定する宛先の上限 |

```bash
gcloud run services update domain-guard --region asia-northeast1 \
  --set-env-vars 'ALLOWED_DOMAINS=example.com,*.contoso.co.jp'
```

### クライアント（`addin/src/config.js`）

| キー | 既定値 | 説明 |
| --- | --- | --- |
| `timeoutMs` | `8000` | API 呼び出しのタイムアウト |
| `failMode` | `'prompt'` | API 失敗時: `prompt`（Outlook 標準の確認）/ `block`（送信中止）/ `allow`（そのまま送信） |
| `maxRecipientsInDialog` | `20` | ダイアログに個別表示する宛先の件数 |
| `apiKey` | `''` | `API_KEY` を設定した場合に指定 |

変更後は `node scripts/build.js <URL>` と再デプロイが必要です。

### 送信時の挙動（`addin/manifest.xml` の `SendMode`）

| 値 | 挙動 |
| --- | --- |
| `SoftBlock`（既定） | 「送信しない」を選ぶと送信が中止される。ハンドラーが動かない環境では送信可能 |
| `Block` | `SoftBlock` に加え、ハンドラーが完了できない場合も送信不可（オフライン時など厳格） |
| `PromptUser` | 独自ダイアログで「送信しない」を選んでも Outlook 標準の「このまま送信」が出てしまうため非推奨 |

## テスト

```bash
npm test                 # サーバー側ロジックのユニットテスト（13 件）
npm start                # http://localhost:8080 でローカル起動（API の確認用）
```

ローカル起動は API 動作確認用です。Outlook にサイドロードするには HTTPS が必要なため、
アドイン自体の確認は Cloud Run にデプロイして行うのが簡単です
（ローカルで行う場合は `npx office-addin-dev-certs install` などで HTTPS 化してください）。

## 注意事項・制限

- **セキュリティ境界ではありません。** ユーザーはアドインを無効化できますし、
  Outlook モバイルや IMAP クライアントからの送信には適用されません。
  誤送信防止の「注意喚起」として設計されています。確実に止めるには
  Exchange のメールフロールール（トランスポートルール）や DLP を併用してください。
- **`API_KEY` はクライアント JS に埋め込まれる**ため、ユーザーからは参照可能です。
  秘匿が必要な場合は `ALLOWED_ORIGINS` の設定に加え、
  Office SSO（`getAccessTokenAsync`）で取得した Microsoft Entra ID トークンを
  サーバーで検証する方式に切り替えてください。
- 宛先が多いメールや API が遅い場合、送信までに数秒待たされます
  （`timeoutMs` 経過後は `failMode` に従います）。
- Cloud Run はコールドスタートがあるため、待ち時間が気になる場合は
  `--min-instances=1` を設定してください。
- 送信時イベントのハンドラーは Outlook 側の制限で一定時間内（数分）に完了する必要があります。
  ダイアログを開いたまま放置すると、送信がキャンセルされる場合があります。
- ログには宛先ドメイン（および件数）が出力されます。メールアドレス本体は出力していませんが、
  運用ポリシーに応じて `server/index.js` の `console.log` を調整してください。
