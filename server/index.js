'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const functions = require('@google-cloud/functions-framework');

const { createAllowlistLoader, domainOf, isAllowedDomain } = require('./lib/allowlist');

const PUBLIC_DIR = path.join(__dirname, 'public');
const MAX_RECIPIENTS = Number(process.env.MAX_RECIPIENTS || 500);
const AUTO_ALLOW_SENDER_DOMAIN = process.env.AUTO_ALLOW_SENDER_DOMAIN !== 'false';
const API_KEY = process.env.API_KEY || '';
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

const allowlist = createAllowlistLoader();

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'text/xml; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

function applyCors(req, res) {
  const origin = req.headers.origin;
  if (!origin) return;
  const ok = ALLOWED_ORIGINS.length === 0 || ALLOWED_ORIGINS.includes(origin);
  if (!ok) return;
  res.set('Access-Control-Allow-Origin', origin);
  res.set('Vary', 'Origin');
  res.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.set('Access-Control-Allow-Headers', 'Content-Type, X-Api-Key');
  res.set('Access-Control-Max-Age', '3600');
}

/** Compares two secrets in constant time (hashing first so lengths don't leak). */
function safeEqual(provided, expected) {
  const a = crypto.createHash('sha256').update(String(provided)).digest();
  const b = crypto.createHash('sha256').update(String(expected)).digest();
  return crypto.timingSafeEqual(a, b);
}

function authorized(req) {
  if (!API_KEY) return true;
  return safeEqual(req.get('x-api-key') || '', API_KEY);
}

/**
 * Admin-only endpoints require ADMIN_API_KEY. It is read at call time and is never
 * embedded in the add-in, unlike API_KEY. When it is unset, admin endpoints don't exist.
 */
function adminAuthorized(req) {
  const adminKey = process.env.ADMIN_API_KEY || '';
  if (!adminKey) return false;
  const provided = req.get('x-admin-key') || '';
  return provided !== '' && safeEqual(provided, adminKey);
}

function notFound(res) {
  res.status(404).send('Not found');
}

/** Normalizes the request body into a flat recipient array. */
function readRecipients(body) {
  const out = [];
  const push = (entry, defaultType) => {
    if (!entry) return;
    if (typeof entry === 'string') {
      out.push({ address: entry, type: defaultType, displayName: '' });
      return;
    }
    out.push({
      address: String(entry.address || entry.emailAddress || ''),
      type: String(entry.type || defaultType || 'to'),
      displayName: String(entry.displayName || ''),
    });
  };

  if (Array.isArray(body.recipients)) body.recipients.forEach((r) => push(r, 'to'));
  ['to', 'cc', 'bcc'].forEach((field) => {
    if (Array.isArray(body[field])) body[field].forEach((r) => push(r, field));
  });
  return out.slice(0, MAX_RECIPIENTS);
}

async function handleCheck(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'method_not_allowed' });
    return;
  }
  if (!authorized(req)) {
    res.status(401).json({ error: 'unauthorized' });
    return;
  }

  const body = typeof req.body === 'object' && req.body ? req.body : {};
  const recipients = readRecipients(body);
  const list = await allowlist.get();

  const extraAllowed = [];
  const senderDomain = domainOf(String(body.sender || body.from || ''));
  if (AUTO_ALLOW_SENDER_DOMAIN && senderDomain) extraAllowed.push(senderDomain);
  const domains = list.domains.concat(extraAllowed);

  const allowed = [];
  const blocked = [];
  const seenBlocked = new Set();

  for (const r of recipients) {
    const domain = domainOf(r.address);
    const item = {
      address: r.address,
      displayName: r.displayName,
      type: r.type,
      domain,
    };
    if (domain && isAllowedDomain(domain, domains)) {
      allowed.push(item);
    } else {
      item.reason = domain ? 'domain_not_in_allowlist' : 'unresolved_address';
      const key = `${r.type}|${r.address.toLowerCase()}`;
      if (!seenBlocked.has(key)) {
        seenBlocked.add(key);
        blocked.push(item);
      }
    }
  }

  const blockedDomains = [...new Set(blocked.map((b) => b.domain).filter(Boolean))].sort();

  console.log(
    JSON.stringify({
      severity: 'INFO',
      event: 'domain_check',
      recipients: recipients.length,
      blocked: blocked.length,
      blockedDomains,
      allowlistSource: list.source,
    })
  );

  res.status(200).json({
    ok: true,
    allowlistVersion: list.version,
    allowlistSource: list.source,
    counts: { total: recipients.length, allowed: allowed.length, blocked: blocked.length },
    blockedDomains,
    allowed,
    blocked,
  });
}

/** Returns the full allowlist. Admin only; responds 404 to everyone else. */
async function handleAllowlist(req, res) {
  if (!adminAuthorized(req)) {
    notFound(res);
    return;
  }
  const list = await allowlist.get({ force: req.query.refresh === '1' });
  res.set('Cache-Control', 'no-store');
  res.status(200).json({
    version: list.version,
    source: list.source,
    count: list.domains.length,
    domains: list.domains,
    staleError: list.staleError,
  });
}

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(req.path === '/' ? '/index.html' : req.path);
  const filePath = path.join(PUBLIC_DIR, urlPath);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.status(403).send('Forbidden');
    return;
  }
  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) {
      notFound(res);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.set('Content-Type', MIME[ext] || 'application/octet-stream');
    res.set('Cache-Control', ext === '.xml' ? 'no-cache' : 'public, max-age=300');
    fs.createReadStream(filePath).pipe(res);
  });
}

functions.http('domainGuard', async (req, res) => {
  applyCors(req, res);
  if (req.method === 'OPTIONS') {
    res.status(204).send('');
    return;
  }

  try {
    if (req.path === '/health' || req.path === '/api/health') {
      res.status(200).json({ ok: true, uptime: process.uptime() });
      return;
    }
    if (req.path === '/api/check') {
      await handleCheck(req, res);
      return;
    }
    if (req.path === '/api/allowlist') {
      await handleAllowlist(req, res);
      return;
    }
    serveStatic(req, res);
  } catch (err) {
    console.error(JSON.stringify({ severity: 'ERROR', message: String(err && err.stack) }));
    if (!res.headersSent) res.status(500).json({ error: 'internal_error' });
  }
});

module.exports = { handleCheck, handleAllowlist };
