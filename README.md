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
アドイン: event.completed({ allowEvent: false, errorMessage: "…" })
   ↓ SendMode="PromptUser"
Outlook 標準の Smart Alerts ダイアログ（対象の宛先一覧を表示）
   ├─「このまま送信」→ 送信される
   └─「送信しない」  → 送信が止まる
```

## 構成

| パス | 内容 |
| --- | --- |
| `addin/manifest.xml` | アドインのマニフェスト（`__BASE_URL__` はビルド時に置換） |
| `addin/src/launchevent/` | 送信時イベントハンドラー（宛先取得 → API で判定 → 結果を Outlook に返す） |
| `addin/src/config.js` | クライアント側設定（API URL、タイムアウト、失敗時の挙動） |
| `server/index.js` | Cloud Run functions 本体（`/api/check`、管理用の `/api/allowlist` と静的配信） |
| `server/lib/allowlist.js` | ドメイン正規化と許可リスト照合ロジック |
| `server/config/allowlist.json` | 同梱の許可ドメイン一覧 |
| `scripts/build.js` | `addin/` を `server/public/` にコピーし URL を埋め込む |
| `scripts/deploy.js` | ビルド → `gcloud run deploy` → URL 確定後に再デプロイ（gcloud 派） |
| `terraform/` | 同じ構成を Terraform で管理する場合（Terraform 派） |

確認ダイアログは Outlook 組み込みの Smart Alerts ダイアログを使います。イベントベースアクティブ化の
ハンドラー内では `displayDialogAsync` などの UI API が
[使用できない](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation#unsupported-apis)ため、
独自ダイアログは実装できません（本文は 500 文字まで、ボタンは「このまま送信 / 送信しない」固定）。

アドインの静的ファイルは同じ Cloud Run サービスから配信します。アドインと API が
**同一オリジン**になるため CORS 設定が不要で、管理する URL も 1 つで済みます。

## 前提条件

- Outlook のアカウント
  - Microsoft 365 の職場・学校アカウント（Exchange Online）: 公式サポート対象
  - 個人の Outlook.com アカウント（Hotmail/live.com 等）: **Outlook on the web で動作を確認済み**（2026-09-13、
    診断ログで `accountType=outlookCom`・Mailbox 1.12 対応を確認）。ただし Smart Alerts の
    [公式サポート表](https://learn.microsoft.com/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events#supported-clients-and-platforms)
    には Exchange Online / Exchange Server しか記載がなく、保証された構成ではありません。
    組織で運用する場合は職場・学校アカウントを前提にしてください
- Outlook 要件セット **Mailbox 1.12** 以上
  - Outlook on the web / 新しい Outlook for Windows: 対応済み
  - classic Outlook for Windows: 2206 (build 15330.20196) 以上
  - Outlook for Mac: 16.65 (22082700) 以上
  - **Outlook モバイル（Android/iOS）は送信時イベント非対応**
- Google Cloud プロジェクト
  - gcloud 派: `gcloud` CLI
  - Terraform 派: Terraform 1.5 以上 と Application Default Credentials（`gcloud auth application-default login`）
- Node.js 20 以上（Terraform 派でも plan 時のビルドに必要）

## 1. デプロイ

### 1-a. gcloud スクリプトでデプロイする場合

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

### 1-b. Terraform でデプロイする場合

`gcloud` スクリプトの代わりに `terraform/` で同じ構成を宣言的に管理できます
（Cloud Run functions、ソース用/許可リスト用の GCS バケット、専用サービスアカウント、公開 IAM）。

```bash
cd terraform
cp terraform.tfvars.example terraform.tfvars   # project_id などを編集
terraform init
terraform plan                                  # この時点で scripts/build.js が実行され server/public と dist/manifest.xml が生成される
terraform apply
```

- **URL の埋め込みは 1 回の apply で完結します。** Cloud Run の決定的 URL
  (`https://<name>-<プロジェクト番号>.<region>.run.app`) を plan 時に計算し、
  `external` データソース経由で `scripts/build.js` を呼んで静的ファイルに埋め込んでから zip →
  アップロード → 関数デプロイ、という順で処理します。したがって `terraform` を実行する環境に **Node.js が必要**です。
- 許可リストは既定で GCS バケット（`<project>-<name>-config/allowlist.json`）に置かれ、
  `server/config/allowlist.json` を編集して `terraform apply` すると関数の再デプロイなしで反映されます
  （最大 5 分のキャッシュあり）。環境変数で指定したい場合は `allowed_domains` 変数を使います。
- ビルド用・実行用に専用のサービスアカウントを作成し、デフォルトの Compute SA は使いません。
  組織ポリシーで `allUsers` への公開が禁止されている場合（`iam.allowedPolicyMemberDomains`）は
  apply が失敗するので、ポリシー側で例外を設定してください。
- `terraform output` の `service_uri` が `base_url` と異なる文字列になることがありますが
  （旧形式の `https://<name>-<hash>-an.a.run.app`）、どちらも同じサービスに届くので問題ありません。
  アドインには決定的 URL の `base_url` が埋め込まれます。
- state は既定でローカルです。チームで運用する場合は `versions.tf` の `backend "gcs"` を有効にしてください。
- `terraform` 実行者には作成した 2 つのサービスアカウントに対する `iam.serviceAccounts.actAs` が必要です（プロジェクトのオーナー/編集者なら可）。

### 動作確認

```bash
curl https://<サービス URL>/health
curl -X POST https://<サービス URL>/api/check \
  -H 'Content-Type: application/json' \
  -d '{"sender":"me@example.com","to":[{"address":"a@example.com"},{"address":"b@gmail.com"}]}'
```

## 2. Outlook への登録

> **登録するのは `dist/manifest.xml` です。** `addin/manifest.xml` は `__BASE_URL__` を含むテンプレートなので、
> そのまま追加すると「IconUrl の値 '__BASE_URL__/…' は正しい形式の URL ではありません」で失敗します。
> 「URL から追加」で `https://<サービス URL>/manifest.xml` を指定しても同じものが使えます。

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

1. Outlook on the web にサインインした状態で（職場・学校アカウント、または個人の Outlook.com）、
   同じブラウザーから **https://aka.ms/olksideload** を開く（「Outlook 用アドイン」ダイアログが直接開きます。
   現在の Outlook on the web の設定画面には「アドインを管理」の項目はありません）
2. 「マイ アドイン」→ 下部の「カスタム アドイン」→「カスタム アドインの追加」→「ファイルから追加...」
3. `dist/manifest.xml` を選択し、警告ダイアログで「インストール」
4. **Outlook のタブを強制リロード（Ctrl+Shift+R）**（イベントハンドラーの登録はリロード後に有効になります）
5. 新規メールで許可リスト外のドメイン宛（例: `someone@gmail.com`）を入れて「送信」→ 確認ダイアログが出れば OK

> 「ファイルから追加」が表示されない場合は、テナント管理者がカスタム アドインのインストールを
> 禁止しています（Microsoft 365 管理センター →「設定」→「組織設定」→「ユーザーが所有するアプリとサービス」、
> または Exchange の OwaMailboxPolicy）。その場合は下の管理者配置を使ってください。
>
> 事前チェック: `npx office-addin-manifest validate dist/manifest.xml` でマニフェストの妥当性を確認できます。

### 組織全体に配布・強制する

1. [Microsoft 365 管理センター](https://admin.microsoft.com/) →「設定」→「統合アプリ」→「カスタム アプリのアップロード」
2. 「App type」で **Office Add-in**（Unified manifest ではなく add-in 専用マニフェスト）を選択
3. `dist/manifest.xml` を選択してアップロード
4. 割り当て先（全員 / 特定のユーザー・グループ）を選択
5. 反映には最大 24 時間かかることがあります（通常は数時間〜）。対象ユーザーは Outlook をリロードしてください

### うまく動かないとき

| 症状 | 原因・対処 |
| --- | --- |
| 「予想以上に時間が掛かっています」が出る | 古い JS がブラウザーにキャッシュされている（静的ファイルは最大 5 分キャッシュ）。**Ctrl+Shift+R で強制リロード**。マニフェストを変えた場合はアドインを削除→再追加 |
| 送信してもダイアログが出ない | リロード忘れ、または宛先が許可リスト内（「許可リストの中身を確認する」の方法で確認） |
| 「宛先ドメインの確認ができませんでした」が出る | API 呼び出しの失敗・タイムアウト（コールドスタート等）。下のログで原因を確認 |

ハンドラーは処理の各段階で `/health?stage=...` を呼ぶ診断機能（`debugBeacon`、**既定は無効**）を持っています。
調査するときは `addin/src/config.js` で `debugBeacon: true` にしてビルド・再デプロイ（と強制リロード）すると、
どこで止まったかを Cloud Run のログで追えます。調査が終わったら `false` に戻してください。

```bash
gcloud logging read \
  'resource.type=cloud_run_revision AND resource.labels.service_name=domain-guard AND httpRequest.requestUrl:*' \
  --project <PROJECT_ID> --freshness=30m --limit=30 \
  --format="table(timestamp,httpRequest.requestMethod,httpRequest.requestUrl.segment(3),httpRequest.status)"
```

正常時は `script-loaded` → `associated` → `ready:...` → `handler-start` → `item:ok` → `recipients:N` →
`POST /api/check` → `checked:blocked=N` → `complete:prompt`（または `complete:allowed`）の順に並びます。
`ready:` にはホスト・プラットフォーム・アカウント種別・Mailbox 1.12 対応可否が記録されます。

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

### 許可リストの中身を確認する

許可リストは外部に公開していません。アドインは `/api/check` に宛先を送り、判定結果（どの宛先がブロック対象か）だけを
受け取るので、許可リスト全体を取得する必要はありません。管理者は次のいずれかで確認します。

- **GCS から直接読む**（プロジェクトの権限が必要）:
  `gcloud storage cat gs://<project>-<name>-config/allowlist.json`
- **リポジトリの `server/config/allowlist.json` を見る**（`terraform apply` で GCS にアップロードされる内容と同じ）
- **管理用 API を使う**: `ADMIN_API_KEY`（Terraform では `admin_api_key` 変数）を設定した場合だけ、
  `curl -H "X-Admin-Key: <キー>" https://<サービス URL>/api/allowlist` で取得できます。
  `?refresh=1` を付けると GCS から即時に再読み込みします（通常は最大 5 分キャッシュ）。
  キーが未設定・不一致のときは 404 を返し、エンドポイントの存在自体を知らせません。
  このキーはサーバー側だけが持ち、アドインの JS には埋め込まれません

> `/api/check` は公開のままなので、宛先を 1 件ずつ送れば「そのドメインが許可されているか」は推測できます。
> 一覧をまとめて取られることは防げますが、1 件ずつの推測まで防ぐには利用者の認証（Entra ID トークンの検証など）が必要です。

## 設定項目

### サーバー（環境変数）

| 変数 | 既定値 | 説明 |
| --- | --- | --- |
| `ALLOWED_DOMAINS` | （なし） | カンマ区切りの許可ドメイン。最優先 |
| `ALLOWLIST_GCS_URI` | （なし） | `gs://bucket/allowlist.json` |
| `ALLOWLIST_CACHE_TTL_MS` | `300000` | 許可リストのキャッシュ時間 |
| `AUTO_ALLOW_SENDER_DOMAIN` | `true` | 送信者自身のドメインを自動的に許可 |
| `API_KEY` | （なし） | 設定すると `X-Api-Key` ヘッダーを要求（アドインの JS に埋め込むため利用者から見える） |
| `ADMIN_API_KEY` | （なし） | 設定すると管理用の `/api/allowlist` が `X-Admin-Key` ヘッダー付きで使える。未設定なら常に 404 |
| `ALLOWED_ORIGINS` | （なし） | CORS 許可オリジン（同一オリジン配信なら不要） |
| `MAX_RECIPIENTS` | `500` | 1 リクエストで判定する宛先の上限 |

```bash
gcloud run services update domain-guard --region asia-northeast1 \
  --set-env-vars 'ALLOWED_DOMAINS=example.com,*.contoso.co.jp'
```

### クライアント（`addin/src/config.js`）

| キー | 既定値 | 説明 |
| --- | --- | --- |
| `timeoutMs` | `3500` | API 呼び出しのタイムアウト。ハンドラー全体の期限（4.5 秒）と Outlook の閾値（5 秒）に収まるようにする |
| `failMode` | `'prompt'` | API 失敗時: `prompt`（確認ダイアログを出す）/ `allow`（そのまま送信） |
| `maxRecipientsInMessage` | `8` | ダイアログ本文に個別表示する宛先の件数（本文は 500 文字まで。超える場合はドメイン一覧に切り替え） |
| `apiKey` | `''` | `API_KEY` を設定した場合に指定 |
| `debugBeacon` | `false` | 処理の各段階で `/health?stage=...` を呼ぶ診断機能。送信 1 回につき 8 リクエスト程度増えるので、調査するときだけ `true` に |

変更後は `node scripts/build.js <URL>` と再デプロイが必要です。

### 送信時の挙動（`addin/manifest.xml` の `SendMode`）

| 値 | 挙動 |
| --- | --- |
| `PromptUser`（既定） | 許可リスト外があると「このまま送信 / 送信しない」の確認が出る。アドインが動かない環境では送信される |
| `SoftBlock` | 「送信しない」のみで、許可リスト外には送れなくなる。アドインが動かない環境では送信される |
| `Block` | `SoftBlock` に加え、アドインが完了できない場合（オフライン等）も送信不可。最も厳格 |

「注意喚起」ではなく「許可リスト外への送信を禁止」したい場合は `SoftBlock` または `Block` に変更してください
（マニフェストの変更後は、サイドロードならアドインの削除→再追加、管理者配置なら再アップロードが必要です）。

## テスト

```bash
npm test                 # サーバー側ロジックのユニットテスト（17 件）
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
- ハンドラーが 5 秒以内に完了しないと Outlook が「予想以上に時間が掛かっています」ダイアログを出します。
  このためハンドラーは **4.5 秒で必ず完了する保険タイマー**を持ち、API の `timeoutMs` は 3.5 秒にしてあります。
  期限を過ぎた場合は `failMode` に従い、既定では「確認できませんでした」の確認ダイアログを出します。
- Cloud Run はコールドスタートがあるため（1〜3 秒程度）、「確認できませんでした」が頻繁に出るようなら
  `min_instances = 1`（Terraform）または `--min-instances=1`（gcloud）を設定してください。
- ログには宛先ドメイン（および件数）が出力されます。メールアドレス本体は出力していませんが、
  運用ポリシーに応じて `server/index.js` の `console.log` を調整してください。
  `debugBeacon` のログにはアドレスは含まれず、段階名と件数のみです。
