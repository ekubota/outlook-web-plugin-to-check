variable "project_id" {
  description = "デプロイ先の Google Cloud プロジェクト ID"
  type        = string
}

variable "region" {
  description = "Cloud Run functions のリージョン"
  type        = string
  default     = "asia-northeast1"
}

variable "name" {
  description = "関数（= Cloud Run サービス）名。URL の一部になる"
  type        = string
  default     = "domain-guard"

  validation {
    condition     = can(regex("^[a-z][a-z0-9-]{0,61}[a-z0-9]$", var.name))
    error_message = "name は小文字英数字とハイフンのみ（先頭は英字）で 63 文字以内にしてください。"
  }
}

variable "base_url" {
  description = <<-EOT
    アドインとマニフェストに埋め込む公開 URL。省略時は Cloud Run の決定的 URL
    (https://<name>-<project番号>.<region>.run.app) を使う。カスタムドメインを
    割り当てる場合などに上書きする。末尾スラッシュなし。
  EOT
  type        = string
  default     = null

  validation {
    condition     = var.base_url == null || can(regex("^https://[^/]+$", var.base_url))
    error_message = "base_url は https:// で始まり、パスや末尾スラッシュを含まない形式にしてください。"
  }
}

# ---- 許可リスト --------------------------------------------------------------

variable "allowed_domains" {
  description = <<-EOT
    許可ドメインを環境変数 ALLOWED_DOMAINS で直接指定する場合のリスト
    （例: ["example.com", "*.contoso.co.jp"]）。指定すると GCS / 同梱 JSON より優先される。
  EOT
  type        = list(string)
  default     = []
}

variable "allowlist_in_gcs" {
  description = "許可リスト JSON を GCS バケットに置き、再デプロイなしで更新できるようにする"
  type        = bool
  default     = true
}

variable "allowlist_file" {
  description = "allowlist_in_gcs=true のとき GCS にアップロードする JSON（terraform/ からの相対パス）"
  type        = string
  default     = "../server/config/allowlist.json"
}

variable "auto_allow_sender_domain" {
  description = "送信者自身のドメインを自動的に許可する"
  type        = bool
  default     = true
}

# ---- API アクセス制御 --------------------------------------------------------

variable "api_key" {
  description = "設定すると API が X-Api-Key ヘッダーを要求する（addin/src/config.js の apiKey にも同じ値を設定すること）"
  type        = string
  default     = ""
  sensitive   = true
}

variable "admin_api_key" {
  description = <<-EOT
    管理用の /api/allowlist を有効にするキー（X-Admin-Key ヘッダーで指定）。
    空のままなら /api/allowlist は常に 404（非公開）。アドインには埋め込まれない。
  EOT
  type        = string
  default     = ""
  sensitive   = true
}

variable "allowed_origins" {
  description = "CORS で許可するオリジン。アドインと API が同一オリジンなら不要"
  type        = list(string)
  default     = []
}

variable "allow_unauthenticated" {
  description = "誰でも呼び出せるようにする（Outlook から呼ぶには true が必要）"
  type        = bool
  default     = true
}

# ---- スケール・リソース ------------------------------------------------------

variable "min_instances" {
  description = "最小インスタンス数。0 なら未使用時は 0 台で費用がかからないが、コールドスタートが発生する（1 以上は常時課金）"
  type        = number
  default     = 0

  validation {
    condition     = var.min_instances >= 0
    error_message = "min_instances は 0 以上にしてください（max_instances 以下であること）。"
  }
}

variable "max_instances" {
  description = "最大インスタンス数。1 ならスケールアウトせず、大量アクセス時も課金に上限がかかる（1 台で同時 80 リクエスト、超過分は 429）"
  type        = number
  default     = 1

  validation {
    condition     = var.max_instances >= 1
    error_message = "max_instances は 1 以上にしてください。"
  }
}

variable "memory" {
  type    = string
  default = "512Mi"
}

variable "cpu" {
  type    = string
  default = "1"
}

variable "timeout_seconds" {
  type    = number
  default = 60
}

variable "labels" {
  description = "全リソースに付与するラベル"
  type        = map(string)
  default     = { app = "outlook-domain-guard" }
}
