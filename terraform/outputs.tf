output "base_url" {
  description = "アドインとマニフェストに埋め込まれた公開 URL"
  value       = local.base_url
}

output "service_uri" {
  description = "Cloud Run が実際に発行した URL（base_url と一致するはず。異なる場合は var.base_url で上書き）"
  value       = google_cloudfunctions2_function.this.service_config[0].uri
}

output "function_url" {
  description = "cloudfunctions.net 形式の URL（同じ関数を指す）"
  value       = google_cloudfunctions2_function.this.url
}

output "manifest_path" {
  description = "Outlook（Microsoft 365 管理センター）に登録するマニフェスト"
  value       = abspath("${path.module}/../${data.external.build.result.manifest}")
}

output "allowlist_gcs_uri" {
  description = "許可リストの GCS パス（allowlist_in_gcs=true のとき）"
  value       = local.allowlist_gcs_uri
}

output "runtime_service_account" {
  value = google_service_account.runtime.email
}

output "source_bucket" {
  value = google_storage_bucket.source.name
}
