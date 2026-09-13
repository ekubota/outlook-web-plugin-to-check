#!/usr/bin/env node
'use strict';

/**
 * addin/ の静的ファイルを server/public/ にコピーし、__BASE_URL__ を実 URL に置換する。
 *
 *   node scripts/build.js https://domain-guard-xxxx.a.run.app
 *   BASE_URL=https://... node scripts/build.js
 *   node scripts/build.js https://... --json   # Terraform の external データソース用（結果を JSON で出力）
 *
 * 置換後の manifest.xml は dist/manifest.xml にも出力される（Outlook への登録用）。
 * 内容が変わらないファイルは書き換えず、出力先に残った不要ファイルは削除する（冪等）。
 */

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'addin');
const OUT = path.join(ROOT, 'server', 'public');
const DIST = path.join(ROOT, 'dist');

const args = process.argv.slice(2);
const jsonMode = args.includes('--json');
const positional = args.filter((a) => !a.startsWith('--'));
const log = jsonMode ? console.error : console.log;

const baseUrl = (positional[0] || process.env.BASE_URL || '').replace(/\/+$/, '');
if (!baseUrl) {
  console.error('BASE_URL を指定してください。例: node scripts/build.js https://domain-guard-xxxx.a.run.app');
  process.exit(1);
}
if (!/^https:\/\//.test(baseUrl) && !/^http:\/\/localhost(:\d+)?$/.test(baseUrl)) {
  console.error('BASE_URL は https:// か http://localhost である必要があります（Outlook は HTTPS 必須）。');
  process.exit(1);
}

const TEXT_EXT = new Set(['.html', '.js', '.css', '.xml', '.json']);
const hash = crypto.createHash('sha256').update(baseUrl);
const written = new Set();

/** 内容が同じなら書き込まない（mtime を無駄に更新しない）。 */
function writeIfChanged(file, content) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (fs.existsSync(file) && fs.readFileSync(file).equals(content)) return;
  fs.writeFileSync(file, content);
}

function copyTree(src, dest, rel = '') {
  for (const entry of fs.readdirSync(src, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
    const from = path.join(src, entry.name);
    const to = path.join(dest, entry.name);
    const relPath = rel ? `${rel}/${entry.name}` : entry.name;
    if (entry.isDirectory()) {
      copyTree(from, to, relPath);
      continue;
    }
    let content = fs.readFileSync(from);
    if (TEXT_EXT.has(path.extname(entry.name).toLowerCase())) {
      content = Buffer.from(content.toString('utf8').split('__BASE_URL__').join(baseUrl));
    }
    hash.update(relPath).update(content);
    writeIfChanged(to, content);
    written.add(path.resolve(to));
  }
}

/** 出力先にある、今回書き出さなかったファイル・空ディレクトリを削除する。 */
function pruneStale(dir) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      pruneStale(p);
      if (fs.readdirSync(p).length === 0) fs.rmdirSync(p);
    } else if (!written.has(path.resolve(p))) {
      fs.unlinkSync(p);
    }
  }
}

copyTree(SRC, OUT);
pruneStale(OUT);

const manifestOut = path.join(DIST, 'manifest.xml');
writeIfChanged(manifestOut, fs.readFileSync(path.join(OUT, 'manifest.xml')));

const result = {
  base_url: baseUrl,
  public_dir: path.relative(ROOT, OUT),
  manifest: path.relative(ROOT, manifestOut),
  hash: hash.digest('hex'),
};

log(`base URL : ${result.base_url}`);
log(`static   : ${result.public_dir}`);
log(`manifest : ${result.manifest}`);
if (jsonMode) process.stdout.write(JSON.stringify(result) + '\n');
