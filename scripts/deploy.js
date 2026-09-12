#!/usr/bin/env node
'use strict';

/**
 * Cloud Run (Cloud Run functions) へデプロイする。
 *
 *   node scripts/deploy.js --project my-proj --region asia-northeast1 [--service domain-guard]
 *   （環境変数 GOOGLE_CLOUD_PROJECT / REGION / SERVICE でも指定可）
 *
 * アドインの manifest とクライアント JS にはサービスの URL を埋め込む必要があるため、
 * 1) URL を推定してビルド → 2) デプロイ → 3) 実際の URL が違えば再ビルドして再デプロイ
 * という順で処理する。gcloud CLI が必要。
 */

const { execFileSync } = require('child_process');
const path = require('path');

const ROOT = path.join(__dirname, '..');

function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

const PROJECT = arg('project', process.env.GOOGLE_CLOUD_PROJECT || '');
const REGION = arg('region', process.env.REGION || 'asia-northeast1');
const SERVICE = arg('service', process.env.SERVICE || 'domain-guard');

if (!PROJECT) {
  console.error('--project か GOOGLE_CLOUD_PROJECT でプロジェクト ID を指定してください。');
  process.exit(1);
}

const gcloudCmd = process.platform === 'win32' ? 'gcloud.cmd' : 'gcloud';

function gcloud(args, opts = {}) {
  return execFileSync(gcloudCmd, args, { encoding: 'utf8', stdio: ['inherit', 'pipe', 'inherit'], ...opts }).trim();
}

function build(baseUrl) {
  execFileSync(process.execPath, [path.join(ROOT, 'scripts', 'build.js'), baseUrl], { stdio: 'inherit' });
}

function deploy() {
  gcloud([
    'run', 'deploy', SERVICE,
    '--source', path.join(ROOT, 'server'),
    '--function', 'domainGuard',
    '--base-image', 'nodejs22',
    '--project', PROJECT,
    '--region', REGION,
    '--allow-unauthenticated',
    '--quiet',
  ], { stdio: 'inherit' });

  return gcloud(['run', 'services', 'describe', SERVICE,
    '--project', PROJECT, '--region', REGION, '--format', 'value(status.url)']);
}

const projectNumber = gcloud(['projects', 'describe', PROJECT, '--format', 'value(projectNumber)']);
let baseUrl = process.env.BASE_URL || `https://${SERVICE}-${projectNumber}.${REGION}.run.app`;

console.log(`\n== 1/3 ビルド (想定 URL: ${baseUrl}) ==`);
build(baseUrl);

console.log('\n== 2/3 デプロイ ==');
let actualUrl = deploy();

if (actualUrl && actualUrl !== baseUrl) {
  console.log(`\n== 3/3 URL が想定と異なるため再ビルド・再デプロイ (${actualUrl}) ==`);
  build(actualUrl);
  actualUrl = deploy();
} else {
  console.log('\n== 3/3 URL 一致。再デプロイ不要 ==');
}

console.log(`\n完了: ${actualUrl}`);
console.log(`Outlook に登録するマニフェスト: ${path.join(ROOT, 'dist', 'manifest.xml')}`);
