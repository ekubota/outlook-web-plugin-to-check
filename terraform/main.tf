# Outlook 宛先ドメインチェック — Cloud Run functions へのデプロイ
#
# 流れ:
#   plan 時に scripts/build.js を実行（external）→ server/ を zip（archive）
#   → GCS にアップロード → Cloud Run functions をソースデプロイ → 公開 IAM
#
# アドインに埋め込む URL は Cloud Run の決定的 URL
#   https://<name>-<project番号>.<region>.run.app
# を使うため、デプロイ前に確定できる（2 回 apply する必要がない）。

locals {
  base_url = trimsuffix(
    coalesce(var.base_url, "https://${var.name}-${data.google_project.this.number}.${var.region}.run.app"),
    "/",
  )

  server_dir = "${path.module}/../server"

  allowlist_gcs_uri = var.allowlist_in_gcs ? "gs://${google_storage_bucket.config[0].name}/${google_storage_bucket_object.allowlist[0].name}" : null

  env = merge(
    {
      AUTO_ALLOW_SENDER_DOMAIN = tostring(var.auto_allow_sender_domain)
    },
    length(var.allowed_domains) > 0 ? { ALLOWED_DOMAINS = join(",", var.allowed_domains) } : {},
    local.allowlist_gcs_uri != null ? { ALLOWLIST_GCS_URI = local.allowlist_gcs_uri } : {},
    length(var.allowed_origins) > 0 ? { ALLOWED_ORIGINS = join(",", var.allowed_origins) } : {},
    var.api_key != "" ? { API_KEY = var.api_key } : {},
  )
}

data "google_project" "this" {
  project_id = var.project_id
}

# ---- API 有効化 ---------------------------------------------------------------

resource "google_project_service" "apis" {
  for_each = toset([
    "cloudfunctions.googleapis.com",
    "run.googleapis.com",
    "cloudbuild.googleapis.com",
    "artifactregistry.googleapis.com",
    "storage.googleapis.com",
    "logging.googleapis.com",
    "iam.googleapis.com",
    "cloudresourcemanager.googleapis.com",
  ])

  service            = each.value
  disable_on_destroy = false
}

# ---- サービスアカウント -------------------------------------------------------
# デフォルトの Compute SA（Editor 相当）を使わず、用途別に最小権限の SA を用意する。

resource "google_service_account" "runtime" {
  account_id   = "${var.name}-run"
  display_name = "${var.name} runtime"
  depends_on   = [google_project_service.apis]
}

resource "google_service_account" "build" {
  account_id   = "${var.name}-build"
  display_name = "${var.name} Cloud Build"
  depends_on   = [google_project_service.apis]
}

# ソースからのビルドに必要な権限（Google 推奨のユーザー管理ビルド SA 構成）
resource "google_project_iam_member" "build" {
  for_each = toset([
    "roles/cloudbuild.builds.builder",
    "roles/artifactregistry.writer",
    "roles/logging.logWriter",
  ])

  project = var.project_id
  role    = each.value
  member  = google_service_account.build.member
}

resource "google_storage_bucket_iam_member" "build_reads_source" {
  bucket = google_storage_bucket.source.name
  role   = "roles/storage.objectViewer"
  member = google_service_account.build.member
}

# IAM の反映待ち。直後に関数を作ると権限エラーで失敗することがある。
resource "time_sleep" "iam_propagation" {
  create_duration = "30s"

  depends_on = [
    google_project_iam_member.build,
    google_storage_bucket_iam_member.build_reads_source,
  ]
}

# ---- ソースのビルドとアップロード ---------------------------------------------

# plan 時に addin/ → server/public/ を生成（__BASE_URL__ を置換）。
# 出力の hash は内容が変わったときだけ変わる（build.js は冪等）。
data "external" "build" {
  program     = ["node", "../scripts/build.js", local.base_url, "--json"]
  working_dir = path.module
}

data "archive_file" "source" {
  type        = "zip"
  source_dir  = local.server_dir
  output_path = "${path.module}/.build/source-${data.external.build.result.hash}.zip"
  excludes    = ["node_modules", "node_modules/**", "test", "test/**"]
}

resource "google_storage_bucket" "source" {
  name                        = "${var.project_id}-${var.name}-source"
  location                    = var.region
  uniform_bucket_level_access = true
  force_destroy               = true
  labels                      = var.labels

  lifecycle_rule {
    condition {
      age = 30
    }
    action {
      type = "Delete"
    }
  }

  depends_on = [google_project_service.apis]
}

# 内容のハッシュを名前に含めることで、変更時だけ新しいオブジェクト → 関数の更新になる。
resource "google_storage_bucket_object" "source" {
  name   = "source-${data.archive_file.source.output_md5}.zip"
  bucket = google_storage_bucket.source.name
  source = data.archive_file.source.output_path
}

# ---- 許可リスト（GCS） --------------------------------------------------------

resource "google_storage_bucket" "config" {
  count = var.allowlist_in_gcs ? 1 : 0

  name                        = "${var.project_id}-${var.name}-config"
  location                    = var.region
  uniform_bucket_level_access = true
  force_destroy               = true
  labels                      = var.labels

  versioning {
    enabled = true
  }

  depends_on = [google_project_service.apis]
}

resource "google_storage_bucket_object" "allowlist" {
  count = var.allowlist_in_gcs ? 1 : 0

  name         = "allowlist.json"
  bucket       = google_storage_bucket.config[0].name
  source       = "${path.module}/${var.allowlist_file}"
  content_type = "application/json"
}

resource "google_storage_bucket_iam_member" "runtime_reads_config" {
  count = var.allowlist_in_gcs ? 1 : 0

  bucket = google_storage_bucket.config[0].name
  role   = "roles/storage.objectViewer"
  member = google_service_account.runtime.member
}

# ---- Cloud Run functions ------------------------------------------------------

resource "google_cloudfunctions2_function" "this" {
  name        = var.name
  location    = var.region
  description = "Outlook add-in: checks recipient domains against an allowlist before send"
  labels      = var.labels

  build_config {
    runtime         = "nodejs22"
    entry_point     = "domainGuard"
    service_account = google_service_account.build.id

    source {
      storage_source {
        bucket = google_storage_bucket.source.name
        object = google_storage_bucket_object.source.name
      }
    }
  }

  service_config {
    service_account_email            = google_service_account.runtime.email
    available_memory                 = var.memory
    available_cpu                    = var.cpu
    timeout_seconds                  = var.timeout_seconds
    min_instance_count               = var.min_instances
    max_instance_count               = var.max_instances
    max_instance_request_concurrency = 80
    ingress_settings                 = "ALLOW_ALL"
    all_traffic_on_latest_revision   = true
    environment_variables            = local.env
  }

  depends_on = [
    google_project_service.apis,
    time_sleep.iam_propagation,
  ]
}

# ---- 公開設定 -----------------------------------------------------------------
# Outlook のクライアントから直接呼ばれるため、未認証アクセスを許可する。
# run.app URL と cloudfunctions.net URL の両方に対して設定する。

resource "google_cloud_run_v2_service_iam_member" "public" {
  count = var.allow_unauthenticated ? 1 : 0

  project  = var.project_id
  location = var.region
  name     = google_cloudfunctions2_function.this.name
  role     = "roles/run.invoker"
  member   = "allUsers"
}

resource "google_cloudfunctions2_function_iam_member" "public" {
  count = var.allow_unauthenticated ? 1 : 0

  project        = var.project_id
  location       = var.region
  cloud_function = google_cloudfunctions2_function.this.name
  role           = "roles/cloudfunctions.invoker"
  member         = "allUsers"
}
