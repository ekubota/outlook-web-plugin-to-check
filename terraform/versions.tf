terraform {
  required_version = ">= 1.5"

  required_providers {
    google = {
      source  = "hashicorp/google"
      version = ">= 6.0, < 8.0"
    }
    archive = {
      source  = "hashicorp/archive"
      version = "~> 2.4"
    }
    external = {
      source  = "hashicorp/external"
      version = "~> 2.3"
    }
    time = {
      source  = "hashicorp/time"
      version = "~> 0.12"
    }
  }

  # 複数人・CI で運用する場合は state を GCS に置く（バケットは事前に作成）。
  # backend "gcs" {
  #   bucket = "my-terraform-state"
  #   prefix = "outlook-domain-guard"
  # }
}

provider "google" {
  project = var.project_id
  region  = var.region
}
