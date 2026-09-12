'use strict';

const fs = require('fs');
const path = require('path');

const DEFAULT_TTL_MS = 5 * 60 * 1000;

/**
 * Normalizes a domain string: lowercase, trimmed, leading "@" and trailing "."
 * removed. Returns "" for anything unusable.
 */
function normalizeDomain(value) {
  if (typeof value !== 'string') return '';
  let d = value.trim().toLowerCase();
  if (!d) return '';
  if (d.startsWith('@')) d = d.slice(1);
  if (d.includes('@')) d = d.slice(d.lastIndexOf('@') + 1);
  d = d.replace(/\.+$/, '');
  // Strip anything that is obviously not a domain (spaces, angle brackets...).
  if (!/^[a-z0-9.*-]+$/.test(d)) return '';
  return d;
}

/**
 * Extracts the domain part of an SMTP address. Returns "" when the address is
 * not a resolvable SMTP address (unresolved recipient, X500/EX address, ...).
 */
function domainOf(address) {
  if (typeof address !== 'string') return '';
  const at = address.lastIndexOf('@');
  if (at < 0) return '';
  return normalizeDomain(address.slice(at + 1));
}

/**
 * An allowlist entry matches when it is
 *   - exactly the domain               ("example.com")
 *   - a wildcard over its subdomains   ("*.example.com" matches "mail.example.com"
 *     but NOT "example.com" -- list the apex separately if you want both)
 */
function matchesEntry(domain, entry) {
  if (!domain || !entry) return false;
  if (entry.startsWith('*.')) {
    const suffix = entry.slice(1); // ".example.com"
    return domain.endsWith(suffix) && domain.length > suffix.length;
  }
  return domain === entry;
}

function isAllowedDomain(domain, entries) {
  return entries.some((entry) => matchesEntry(domain, entry));
}

function parseEntries(raw) {
  let list = [];
  if (Array.isArray(raw)) {
    list = raw;
  } else if (raw && Array.isArray(raw.allowedDomains)) {
    list = raw.allowedDomains;
  } else if (typeof raw === 'string') {
    list = raw.split(/[,\s\r\n]+/);
  }
  const seen = new Set();
  for (const item of list) {
    const d = normalizeDomain(item);
    if (d) seen.add(d);
  }
  return [...seen].sort();
}

async function loadFromGcs(uri) {
  const match = /^gs:\/\/([^/]+)\/(.+)$/.exec(uri);
  if (!match) throw new Error(`Invalid ALLOWLIST_GCS_URI: ${uri}`);
  const { Storage } = require('@google-cloud/storage');
  const [buf] = await new Storage().bucket(match[1]).file(match[2]).download();
  return JSON.parse(buf.toString('utf8'));
}

function loadFromFile(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

/**
 * Loads the allowlist from, in order of precedence:
 *   1. ALLOWED_DOMAINS            (comma/space separated env var)
 *   2. ALLOWLIST_GCS_URI          (gs://bucket/allowlist.json, cached)
 *   3. config/allowlist.json      (bundled with the deployment)
 * The result is cached in memory for ALLOWLIST_CACHE_TTL_MS.
 */
function createAllowlistLoader(env = process.env) {
  const ttlMs = Number(env.ALLOWLIST_CACHE_TTL_MS || DEFAULT_TTL_MS);
  const localFile = env.ALLOWLIST_FILE || path.join(__dirname, '..', 'config', 'allowlist.json');
  let cache = null;

  async function read() {
    if (env.ALLOWED_DOMAINS) {
      return { source: 'env:ALLOWED_DOMAINS', raw: env.ALLOWED_DOMAINS };
    }
    if (env.ALLOWLIST_GCS_URI) {
      return { source: env.ALLOWLIST_GCS_URI, raw: await loadFromGcs(env.ALLOWLIST_GCS_URI) };
    }
    return { source: `file:${path.basename(localFile)}`, raw: loadFromFile(localFile) };
  }

  return {
    async get({ force = false } = {}) {
      if (!force && cache && Date.now() - cache.loadedAt < ttlMs) return cache;
      try {
        const { source, raw } = await read();
        cache = {
          domains: parseEntries(raw),
          source,
          version: (raw && raw.version) || new Date().toISOString().slice(0, 19) + 'Z',
          loadedAt: Date.now(),
        };
      } catch (err) {
        if (cache) {
          // Serve the stale list rather than failing closed on a transient error.
          cache.staleError = String(err.message || err);
          return cache;
        }
        throw err;
      }
      return cache;
    },
    invalidate() {
      cache = null;
    },
  };
}

module.exports = {
  createAllowlistLoader,
  domainOf,
  isAllowedDomain,
  matchesEntry,
  normalizeDomain,
  parseEntries,
};
