#!/usr/bin/env node
'use strict';

/**
 * 依存パッケージなしでアドイン用アイコン PNG を生成する。
 * 図柄: 角丸の四角 (Outlook ブルー) + 白い「!」マーク。
 *   node scripts/gen-icons.js
 */

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const SIZES = [16, 32, 64, 80, 128];
const OUT_DIR = path.join(__dirname, '..', 'addin', 'assets');
const BG = [15, 108, 189, 255]; // #0f6cbd
const FG = [255, 255, 255, 255];

function crc32(buf) {
  let c;
  const table = crc32.table || (crc32.table = (() => {
    const t = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c;
    }
    return t;
  })());
  let crc = -1;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ -1) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function encodePng(size, pixelAt) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let p = 0;
  for (let y = 0; y < size; y++) {
    raw[p++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixelAt(x, y);
      raw[p++] = r;
      raw[p++] = g;
      raw[p++] = b;
      raw[p++] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** 角丸四角の内側か（アンチエイリアスなしの単純判定）。 */
function insideRoundedRect(x, y, size) {
  const r = size * 0.22;
  const min = size * 0.04;
  const max = size - min;
  const cx = Math.min(Math.max(x + 0.5, min + r), max - r);
  const cy = Math.min(Math.max(y + 0.5, min + r), max - r);
  const dx = x + 0.5 - cx;
  const dy = y + 0.5 - cy;
  return dx * dx + dy * dy <= r * r;
}

/** 「!」の形（縦棒 + 点）。 */
function insideBang(x, y, size) {
  const px = (x + 0.5) / size;
  const py = (y + 0.5) / size;
  const inX = Math.abs(px - 0.5) <= 0.075;
  const bar = inX && py >= 0.24 && py <= 0.62;
  const dot = Math.hypot(px - 0.5, py - 0.745) <= 0.085;
  return bar || dot;
}

fs.mkdirSync(OUT_DIR, { recursive: true });
for (const size of SIZES) {
  const png = encodePng(size, (x, y) => {
    if (!insideRoundedRect(x, y, size)) return [0, 0, 0, 0];
    return insideBang(x, y, size) ? FG : BG;
  });
  const file = path.join(OUT_DIR, `icon-${size}.png`);
  fs.writeFileSync(file, png);
  console.log(`wrote ${path.relative(process.cwd(), file)} (${png.length} bytes)`);
}
