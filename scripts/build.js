#!/usr/bin/env node
'use strict';

/**
 * addin/ の静的ファイルを server/public/ にコピーし、__BASE_URL__ を実 URL に置換する。
 *
 *   node scripts/build.js https://domain-guard-xxxx.a.run.app
 *   BASE_URL=https://... node scripts/build.js
 *
 * 置換後の manifest.xml は dist/manifest.xml にも出力される（Outlook への登録用）。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'addin');
const OUT = path.join(ROOT, 'server', 'public');
const DIST = path.join(ROOT, 'dist');

const baseUrl = (process.argv[2] || process.env.BASE_URL || '').replace(/\/+$/, '');
if (!baseUrl) {
  console.error('BASE_URL を指定してください。例: node scripts/build.js https://domain-guard-xxxx.a.run.app');
  process.exit(1);
}
if (!/^https:\/\//.test(baseUrl) && !/^http:\/\/localhost(:\d+)?$/.test(baseUrl)) {
  console.error('BASE_URL は https:// か http://localhost である必要があります（Outlook は HTTPS 必須）。');
  process.exit(1);
}

const TEXT_EXT = new Set(['.html', '.js', '.css', '.xml', '.json']);

function copyTree(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const from = path.join(src, entry.name);
    const to = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyTree(from, to);
    } else if (TEXT_EXT.has(path.extname(entry.name).toLowerCase())) {
      const body = fs.readFileSync(from, 'utf8').split('__BASE_URL__').join(baseUrl);
      fs.writeFileSync(to, body);
    } else {
      fs.copyFileSync(from, to);
    }
  }
}

fs.rmSync(OUT, { recursive: true, force: true });
copyTree(SRC, OUT);

fs.mkdirSync(DIST, { recursive: true });
fs.copyFileSync(path.join(OUT, 'manifest.xml'), path.join(DIST, 'manifest.xml'));

console.log(`base URL : ${baseUrl}`);
console.log(`static   : ${path.relative(ROOT, OUT)}`);
console.log(`manifest : ${path.relative(ROOT, path.join(DIST, 'manifest.xml'))}`);
