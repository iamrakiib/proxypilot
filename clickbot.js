// ProxyPilot — AdsPower profiles (CREATE → open ALL links in ONE group in PARALLEL → STOP → DELETE)
//
// Behaviour:
//   ✓ Proxies taken sequentially (wraps around when list exhausted)
//   ✓ User-agent picked randomly per profile
//   ✓ IPQS check before each proxy (multi-key rotation on daily limit)
//   ✓ Groups cycle one-per-profile: A → B → C → … → A → B …
//   ✓ All links in a group open simultaneously (parallel tabs)
//   ✓ CTA click → human-like scroll → dwell 30-40s → close profile
//   ✓ Desktop fingerprint (Windows/Linux)
//
// Optional env:
//   ADSP_USE=1
//   ADSP_OS=windows|linux
//   ADSP_BASEURL=http://...
//   ADSP_API_KEY=...
//   ADSP_GROUP_ID=...
//   ADSP_GROUP_NAME=AutoGroup
//   ADSP_GROUPS_PER_PROFILE=A,B,C   (limit which groups cycle; default = all)
//   ADSP_GROUP_MAX_TABS=0
//   ADSP_HEADFUL=1
//   FORCE_DESKTOP=1
//
// Google Search mode:
//   --google-search               (or GOOGLE_SEARCH=1)
//   GOOGLE_SEARCH_QUERY="keyword" (default: example site)

const fs = require('fs');
const path = require('path');
const os = require('os');
const readline = require('readline');
const { chromium } = require('playwright');
const xlsx = require('xlsx');
const http = require('http');
const https = require('https');
const { URL } = require('url');

/* ---------------------- Files ---------------------- */
const GROUPS_FILE = path.join(__dirname, 'groups.json');
const PROFILES_XLSX = path.join(__dirname, 'profiles.xlsx');
const PROXIES_FILE = path.join(__dirname, 'proxies.txt');
const UAS_FILE = path.join(__dirname, 'useragents.txt');
const DESKTOP_UAS_FILE = path.join(__dirname, 'desktop_useragents.txt');
const GROUP_ID_CACHE = path.join(__dirname, '.adsp_group_id.txt');

/* ------------------- Behaviour --------------------- */
const HEADFUL = true;
const DWELL_MS = { min: 4000, max: 8000 };   // 4-8s on page
const HUMAN_SCROLL_MS = {
  min: parseInt(process.env.HUMAN_SCROLL_MIN_MS || '8000', 10),
  max: parseInt(process.env.HUMAN_SCROLL_MAX_MS || '15000', 10),
};
const INTERNAL_LINK_SCROLL_MS = {
  min: parseInt(process.env.INTERNAL_LINK_SCROLL_MIN_MS || '12000', 10),
  max: parseInt(process.env.INTERNAL_LINK_SCROLL_MAX_MS || '16000', 10),
};
const PRENAV_PAUSE_MS = { min: 800, max: 1500 };
const RETRY_NAVIGATION = 1;
const INTER_RUN_SLEEP_MS = { min: 1200, max: 2500 };
const PROFILE_TIMEOUT_MS = 600000; // 10 min hard cap
const GROUP_MAX_TABS = parseInt(process.env.ADSP_GROUP_MAX_TABS || '0', 10);
const SITE_TEST_USE = process.argv.includes('--site-test') || process.env.SITE_TEST === '1';
const MOCK_SEARCH_USE = process.argv.includes('--mock-search-test') || process.env.MOCK_SEARCH_TEST === '1';
const SITE_TEST_URL = process.env.SITE_TEST_URL || '';
const MOCK_SEARCH_QUERY = process.env.MOCK_SEARCH_QUERY || 'example site';
const MOCK_SEARCH_EXACT_NAME = process.env.MOCK_SEARCH_EXACT_NAME || 'example site';
const MOCK_SEARCH_KEEP_OPEN_MS = parseInt(process.env.MOCK_SEARCH_KEEP_OPEN_MS || '0', 10);
const SITE_TEST_SCROLL_MS = {
  min: parseInt(process.env.SITE_TEST_SCROLL_MIN_MS || '30000', 10),
  max: parseInt(process.env.SITE_TEST_SCROLL_MAX_MS || '40000', 10),
};

// ----- Google Search mode flags -----
const GOOGLE_SEARCH_USE = process.argv.includes('--google-search') || process.env.GOOGLE_SEARCH === '1';
const GOOGLE_SEARCH_QUERY = process.env.GOOGLE_SEARCH_QUERY || 'example site';

const FORCE_DESKTOP = (process.env.FORCE_DESKTOP || '1') === '1';
const DESKTOP_VIEWPORT = {
  width: parseInt(process.env.DESKTOP_VIEWPORT_WIDTH || '1366', 10),
  height: parseInt(process.env.DESKTOP_VIEWPORT_HEIGHT || '768', 10),
};

const MOBILE_SIZES = [
  { width: 375, height: 667 },  // iPhone SE
  { width: 390, height: 844 },  // iPhone 14
  { width: 393, height: 851 },  // Pixel 5
  { width: 412, height: 915 },  // Galaxy S21+
  { width: 360, height: 800 },  // Galaxy S21
  { width: 414, height: 896 },  // iPhone 11
  { width: 428, height: 926 },  // iPhone 14 Plus
  { width: 360, height: 780 },  // Galaxy A52
  { width: 412, height: 869 },  // Pixel 4
  { width: 384, height: 854 },  // Galaxy S10e
  { width: 375, height: 812 },  // iPhone X / XS
  { width: 414, height: 736 },  // iPhone 8 Plus
];

/* ---------------- AdsPower config ------------------ */
function readAdsConfig() {
  try {
    const p = path.join(__dirname, 'ads-config.json');
    if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (_) { }
  return {};
}
const ADSCFG = readAdsConfig();
const ADSP_USE = process.argv.includes('--adspower') || process.env.ADSP_USE === '1';
let ADSP_BASE = process.env.ADSP_BASEURL || ADSCFG.baseURL || 'http://local.adspower.net:50325';
const ADSP_APIKEY = process.env.ADSP_API_KEY || ADSCFG.apiKey || '';
const ADSP_HEADFUL = (process.env.ADSP_HEADFUL || (ADSCFG.headful ? '1' : '')).toString() === '1';
const ADSP_GROUP_ID_ENV = process.env.ADSP_GROUP_ID || '';
const ADSP_GROUP_NAME = process.env.ADSP_GROUP_NAME || 'AutoGroup';
const GROUPS_PER_PROFILE = (process.env.ADSP_GROUPS_PER_PROFILE || '')
  .split(',').map(s => s.trim()).filter(Boolean);

// IPQS API keys
let IPQS_KEYS = (() => {
  const envVal = process.env.IPQS_API_KEY || '';
  const cfgVal = ADSCFG.ipqsApiKey || '';
  const fromEnv = envVal ? envVal.split(',').map(s => s.trim()).filter(Boolean) : [];
  const fromCfg = Array.isArray(cfgVal)
    ? cfgVal.filter(Boolean)
    : cfgVal ? cfgVal.split(',').map(s => s.trim()).filter(Boolean) : [];
  return fromEnv.length ? fromEnv : fromCfg;
})();
let ipqsKeyIndex = 0;
const IPQS_STRICTNESS = Number.isInteger(ADSCFG.ipqsStrictness) ? ADSCFG.ipqsStrictness : 1;
const IPQS_MAX_SCORE = Number.isFinite(ADSCFG.ipqsMaxScore) ? ADSCFG.ipqsMaxScore : 30;

/* -------- IPQS daily exhaustion persistence --------- */
const IPQS_USAGE_FILE = path.join(__dirname, '.ipqs_daily.json');

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function loadExhaustedKeys() {
  try {
    if (fs.existsSync(IPQS_USAGE_FILE)) {
      const data = JSON.parse(fs.readFileSync(IPQS_USAGE_FILE, 'utf8'));
      if (data.date === todayStr()) return new Set(data.exhausted || []);
      fs.writeFileSync(IPQS_USAGE_FILE, JSON.stringify({ date: todayStr(), exhausted: [] }), 'utf8');
    }
  } catch { }
  return new Set();
}

function markKeyExhausted(key) {
  try {
    let data = { date: todayStr(), exhausted: [] };
    if (fs.existsSync(IPQS_USAGE_FILE)) {
      try {
        const d = JSON.parse(fs.readFileSync(IPQS_USAGE_FILE, 'utf8'));
        if (d.date === todayStr()) data = d;
      } catch { }
    }
    if (!data.exhausted.includes(key)) {
      data.exhausted.push(key);
      fs.writeFileSync(IPQS_USAGE_FILE, JSON.stringify(data), 'utf8');
    }
  } catch (e) { console.warn('[IPQS] Could not save exhausted key:', e.message); }
}

const ADSP_OS = (process.env.ADSP_OS || (
  process.platform === 'win32' ? 'windows' :
    process.platform === 'linux' ? 'linux' :
      'windows'
)).toLowerCase();

/* ----------------- Helpers ------------------------- */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function rint(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }
function jitter(min, max) { return rint(min, max); }
function choose(arr) { return (!arr || arr.length === 0) ? null : arr[Math.floor(Math.random() * arr.length)]; }
function notEmpty(x) { return !!x && String(x).trim().length > 0; }

/* ========== CAPTCHA Solver (fallback) ========== */
const TWOCAPTCHA_API_KEY = process.env.TWOCAPTCHA_API_KEY || '';
if (!TWOCAPTCHA_API_KEY) {
  console.warn('[WARN] TWOCAPTCHA_API_KEY not set – bot solver will fail if extension does not solve.');
}

async function call2Captcha(siteKey, pageUrl, dataS) {
  if (!TWOCAPTCHA_API_KEY) throw new Error('TWOCAPTCHA_API_KEY not set');
  const params = new URLSearchParams({
    key: TWOCAPTCHA_API_KEY,
    method: 'userrecaptcha',
    googlekey: siteKey,
    pageurl: pageUrl,
    json: '1'
  });
  if (dataS) params.set('data-s', dataS);

  const submitUrl = `https://2captcha.com/in.php?${params.toString()}`;
  console.log('[BotSolver] Submitting CAPTCHA...');
  const submitRes = await fetch(submitUrl);
  const submitData = await submitRes.json();
  if (submitData.status !== 1) {
    throw new Error(`2Captcha submit error: ${submitData.request}`);
  }
  const taskId = submitData.request;

  for (let attempt = 0; attempt < 24; attempt++) {
    await sleep(5000);
    const pollUrl = `https://2captcha.com/res.php?key=${TWOCAPTCHA_API_KEY}&action=get&id=${taskId}&json=1`;
    const pollRes = await fetch(pollUrl);
    const pollData = await pollRes.json();
    if (pollData.status === 1) return pollData.request;
    if (pollData.request !== 'CAPCHA_NOT_READY') {
      throw new Error(`2Captcha poll error: ${pollData.request}`);
    }
  }
  throw new Error('2Captcha timeout');
}

async function extractSiteKeyAndDataS(page) {
  const html = await page.evaluate(() => document.documentElement.outerHTML);
  let siteKey = null, dataS = null;
  const siteKeyMatch = html.match(/data-sitekey=["']([^"']+)["']/i) || html.match(/render=([^&"']+)/i);
  if (siteKeyMatch) siteKey = siteKeyMatch[1];
  const dataSMatch = html.match(/data-s["']?\s*[:=]\s*["']([^"']+)["']/i) ||
    html.match(/data-s=(["'])([^"']+)\1/i) ||
    html.match(/data_s\s*[:=]\s*["']([^"']+)["']/i) ||
    html.match(/["']data-s["']\s*[:=]\s*["']([^"']+)["']/i);
  if (dataSMatch) dataS = dataSMatch[1] || dataSMatch[2];
  return { siteKey, dataS };
}

async function solveGoogleCaptcha(page) {
  console.log('[BotSolver] Extracting siteKey and data-s...');
  const { siteKey, dataS } = await extractSiteKeyAndDataS(page);
  if (!siteKey) throw new Error('Could not find siteKey');
  const pageUrl = page.url();
  console.log(`[BotSolver] Solving with siteKey: ${siteKey.substring(0, 10)}..., dataS: ${dataS ? dataS.substring(0, 10) + '...' : 'none'}`);
  const token = await call2Captcha(siteKey, pageUrl, dataS);
  console.log('[BotSolver] Token received, injecting...');
  await page.evaluate((token) => {
    const field = document.querySelector('textarea[name="g-recaptcha-response"]');
    if (field) { field.value = token; field.dispatchEvent(new Event('change', { bubbles: true })); }
    const callbackEl = document.querySelector('[data-callback]');
    if (callbackEl) {
      const cbName = callbackEl.getAttribute('data-callback');
      if (cbName && typeof window[cbName] === 'function') window[cbName](token);
    }
    const form = document.querySelector('form[action*="/sorry/index"]');
    if (form && form.requestSubmit) form.requestSubmit();
  }, token);
  await page.waitForNavigation({ timeout: 60000 }).catch(() => { });
  console.log('[BotSolver] CAPTCHA solved and page reloaded.');
}
/* ==================== End CAPTCHA Solver ==================== */

/* HTTP JSON (AdsPower) with 127.0.0.1 fallback */
async function httpJson(method, urlString, bodyObj, attempt = 0) {
  return await new Promise((resolve, reject) => {
    const u = new URL(urlString);
    const isHttps = u.protocol === 'https:';
    const opts = {
      method,
      hostname: u.hostname,
      port: u.port || (isHttps ? 443 : 80),
      path: u.pathname + (u.search || ''),
      headers: { 'Content-Type': 'application/json' }
    };
    if (ADSP_APIKEY) {
      opts.headers['X-ADSPOWER-API-KEY'] = ADSP_APIKEY;
      opts.headers['Authorization'] = `Bearer ${ADSP_APIKEY}`;
    }
    const req = (isHttps ? https : http).request(opts, res => {
      let data = ''; res.setEncoding('utf8');
      res.on('data', d => data += d);
      res.on('end', () => {
        try { resolve(JSON.parse(data || '{}')); }
        catch (e) { reject(new Error('Invalid JSON from AdsPower: ' + e.message + ' body=' + data)); }
      });
    });
    req.on('error', async (err) => {
      if ((err.code === 'ENOTFOUND' || err.code === 'EAI_AGAIN') && attempt === 0 && u.hostname === 'local.adspower.net') {
        const alt = urlString.replace('local.adspower.net', '127.0.0.1');
        console.log('[AdsPower] DNS failed; retrying via 127.0.0.1 …');
        ADSP_BASE = ADSP_BASE.replace('local.adspower.net', '127.0.0.1');
        try { const r = await httpJson(method, alt, bodyObj, 1); resolve(r); return; } catch (e) { reject(e); return; }
      }
      reject(err);
    });
    if (bodyObj) req.write(JSON.stringify(bodyObj));
    req.end();
  });
}

/* ----------------- Loaders ------------------------- */
async function ask(q) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(r => rl.question(q, a => { rl.close(); r(a); }));
}
function loadFromExcel(xlsxPath) {
  if (!fs.existsSync(xlsxPath)) return { proxies: [], uas: [], desktopUas: [] };
  const wb = xlsx.readFile(xlsxPath);
  const proxies = [], uas = [], desktopUas = [];
  const findSheet = (...names) => wb.SheetNames.find(n => names.map(s => s.toLowerCase()).includes(n.toLowerCase()));
  const readCol = (sheetName, out) => {
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
    for (const r of rows) { const v = r?.[0]?.toString().trim(); if (v) out.push(v); }
  };
  const sProxies = findSheet('Proxies', 'proxy', 'PROXIES');
  if (sProxies) readCol(sProxies, proxies);
  const sUas = findSheet('UserAgents', 'useragents', 'UA', 'us');
  if (sUas) readCol(sUas, uas);
  const sDesktop = findSheet('Desktop UA', 'DesktopUA', 'Desktop UserAgents', 'Desktop', 'desktopuseragents', 'desktop_ua');
  if (sDesktop) readCol(sDesktop, desktopUas);
  return { proxies, uas, desktopUas };
}
function loadLines(filePath) {
  if (!fs.existsSync(filePath)) return [];
  return fs.readFileSync(filePath, 'utf8').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
}
function loadGroups(file) {
  if (!fs.existsSync(file)) throw new Error('Missing groups.json');
  const raw = JSON.parse(fs.readFileSync(file, 'utf8'));
  const order = Object.keys(raw);
  return { order, groups: raw };
}

/* --------------- Proxy parsing --------------------- */
function parseProxyLine(line) {
  if (!line) return null;
  line = line.trim();
  let scheme = 'http';
  if (/^socks5:\/\//i.test(line)) scheme = 'socks5';
  else if (/^socks4:\/\//i.test(line)) scheme = 'socks4';
  else if (/^https:\/\//i.test(line)) scheme = 'https';
  const authority = line.replace(/^(https?|socks[45]?):\/\//i, '');
  if (authority.includes('@')) {
    const atIdx = authority.lastIndexOf('@');
    const creds = authority.substring(0, atIdx);
    const hostPort = authority.substring(atIdx + 1);
    const lastColon = hostPort.lastIndexOf(':');
    if (lastColon !== -1) {
      const host = hostPort.substring(0, lastColon);
      const port = hostPort.substring(lastColon + 1);
      if (host && /^\d+$/.test(port)) {
        const colonIdx = creds.indexOf(':');
        const username = colonIdx !== -1 ? creds.substring(0, colonIdx) : creds;
        const password = colonIdx !== -1 ? creds.substring(colonIdx + 1) : '';
        return { server: `${scheme}://${host}:${port}`, username, password };
      }
    }
  }
  const parts = authority.split(':');
  if (parts.length === 4 && /^\d+$/.test(parts[3]) && parts[2].includes('.')) {
    const [username, password, host, port] = parts;
    return { server: `${scheme}://${host}:${port}`, username, password };
  }
  if (parts.length === 2 && /^\d+$/.test(parts[1])) {
    return { server: `${scheme}://${parts[0]}:${parts[1]}` };
  }
  console.warn('Could not parse proxy entry; check the supported formats.');
  return null;
}
function getProxyHost(proxyObj) {
  try { return new URL(proxyObj.server).hostname; } catch { return proxyObj.server; }
}

/* ===================== IPQS ======================== */
async function rawCheckIPQS(ip, apiKey) {
  const endpoint = `https://ipqualityscore.com/api/json/ip/${apiKey}/${encodeURIComponent(ip)}?strictness=${IPQS_STRICTNESS}&allow_public_access_points=true&fast=false&lighter_penalties=false&mobile=false`;
  return new Promise((resolve) => {
    const req = https.get(endpoint, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', d => data += d);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.success === true) {
            resolve({ score: json.fraud_score, rateLimited: false });
          } else {
            const msg = (json.message || '').toLowerCase();
            const rateLimited = /limit|quota|exceed|too many|upgrade|insufficient|credits/i.test(msg);
            if (rateLimited) {
              console.warn(`[IPQS] Key limit reached for ${ip}: ${json.message}`);
            } else {
              console.warn(`[IPQS] API error for ${ip}: ${json.message || data}`);
            }
            resolve({ score: null, rateLimited });
          }
        } catch { resolve({ score: null, rateLimited: false }); }
      });
    });
    req.on('error', (e) => {
      console.warn(`[IPQS] Request error for ${ip}:`, e.message);
      resolve({ score: null, rateLimited: false });
    });
    req.setTimeout(12000, () => {
      req.destroy();
      console.warn(`[IPQS] Timeout for ${ip}`);
      resolve({ score: null, rateLimited: false });
    });
  });
}

async function checkWithKeyRotation(ip) {
  if (!IPQS_KEYS.length) return null;
  while (IPQS_KEYS.length > 0) {
    const idx = ipqsKeyIndex % IPQS_KEYS.length;
    const key = IPQS_KEYS[idx];
    const { score, rateLimited } = await rawCheckIPQS(ip, key);
    if (rateLimited) {
      IPQS_KEYS.splice(idx, 1);
      markKeyExhausted(key);
      console.warn(`[IPQS] Key exhausted — saved to disk, ${IPQS_KEYS.length} key(s) remaining today`);
      continue;
    }
    return score;
  }
  console.error('[IPQS] All API keys exhausted for today. Run will stop — restart tomorrow when keys reset.');
  return null;
}

const IP_CHECK_ENDPOINTS = [
  { url: 'http://api.ipify.org/?format=json', host: 'api.ipify.org' },
  { url: 'http://checkip.amazonaws.com/', host: 'checkip.amazonaws.com' },
  { url: 'http://icanhazip.com/', host: 'icanhazip.com' },
  { url: 'http://ip-api.com/json', host: 'ip-api.com' },
  { url: 'http://ifconfig.me/ip', host: 'ifconfig.me' },
];

function parseIPFromBody(data) {
  const text = (data || '').trim();
  try {
    const json = JSON.parse(text);
    const ip = json.ip || json.query || json.origin;
    if (ip && /^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(String(ip).trim())) return String(ip).trim();
  } catch { }
  if (/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(text)) return text;
  return null;
}

async function getExitIPViaProxy(proxyObj) {
  try {
    const proxyUrl = new URL(proxyObj.server);
    if (/^socks/i.test(proxyUrl.protocol)) return null;
    const proxyPort = parseInt(proxyUrl.port) || 80;
    const authHeader = proxyObj.username
      ? `Basic ${Buffer.from(`${proxyObj.username}:${proxyObj.password || ''}`).toString('base64')}`
      : null;
    const makeRequest = (ep) => new Promise((resolve) => {
      try {
        const opts = {
          method: 'GET',
          hostname: proxyUrl.hostname,
          port: proxyPort,
          path: ep.url,
          headers: { 'Host': ep.host, 'Connection': 'close' }
        };
        if (authHeader) opts.headers['Proxy-Authorization'] = authHeader;
        const req = http.request(opts, (res) => {
          let data = '';
          res.setEncoding('utf8');
          res.on('data', d => data += d);
          res.on('end', () => resolve(parseIPFromBody(data)));
        });
        req.on('error', () => resolve(null));
        req.setTimeout(8000, () => { req.destroy(); resolve(null); });
        req.end();
      } catch { resolve(null); }
    });
    return await new Promise((resolve) => {
      let done = 0;
      const total = IP_CHECK_ENDPOINTS.length;
      IP_CHECK_ENDPOINTS.forEach(ep => makeRequest(ep).then(ip => {
        done++;
        if (ip) resolve(ip);
        else if (done === total) resolve(null);
      }));
    });
  } catch { }
  return null;
}

/* ---------------- CTA helpers --------------------- */
const BUTTON_TERMS = [
  "follow link", "continue", "Get Reward", "Check Eligibility Now", "Unlock Access", "click here",
  "Enter Now", "open", "next", "go to", "get started", "proceed", "go", "apply now", "Dating Now",
  "allow", "confirm", "yes", "agree", "visit", "start Now", "View My Options", "Claim"
];
const FALLBACK_SELECTORS = [
  'button.cta', 'button.primary', 'a.cta', 'a.btn', 'button[role="button"]',
  'div[role="button"]', '.btn', '.primary', '.continue', '.next', '.proceed'
];

async function clickSmart(locator, page) {
  const count = await locator.count();
  if (!count) return false;
  const el = locator.first();
  const isMobile = await page.evaluate(() =>
    /Android|iPhone|Mobile/i.test(navigator.userAgent) || window.innerWidth <= 640
  );
  try { await el.waitFor({ state: 'attached', timeout: 2000 }); } catch { }
  try { await el.evaluate(n => n.scrollIntoView({ block: 'center', inline: 'center' })); } catch { }
  try { await page.waitForTimeout(120); } catch { }
  if (isMobile) {
    try { await el.tap({ timeout: 1800 }); return true; } catch { }
  }
  try { await el.click({ timeout: 1800, trial: true }); await el.click({ timeout: 1800 }); return true; } catch { }
  try {
    const box = await el.boundingBox();
    if (box) {
      const cx = Math.round(box.x + box.width / 2);
      const cy = Math.round(box.y + Math.min(box.height * 0.55, box.height - 4));
      if (page.touchscreen && typeof page.touchscreen.tap === 'function') {
        await page.touchscreen.tap(cx, cy);
      } else {
        await page.mouse.click(cx, cy, { delay: 30 });
      }
      return true;
    }
  } catch { }
  try { await el.evaluate(n => n.click()); return true; } catch { }
  return false;
}

async function clickFollowIfExists(page) {
  const tryLoc = loc => clickSmart(loc, page);
  for (const term of BUTTON_TERMS) {
    if (await tryLoc(page.getByRole('button', { name: new RegExp(`\\b${term}\\b`, 'i') }))) return true;
    if (await tryLoc(page.getByRole('link', { name: new RegExp(`\\b${term}\\b`, 'i') }))) return true;
  }
  for (const sel of FALLBACK_SELECTORS) if (await tryLoc(page.locator(sel))) return true;
  for (const f of page.frames()) {
    for (const term of BUTTON_TERMS) {
      if (await tryLoc(f.getByRole('button', { name: new RegExp(`\\b${term}\\b`, 'i') }))) return true;
      if (await tryLoc(f.getByRole('link', { name: new RegExp(`\\b${term}\\b`, 'i') }))) return true;
    }
    for (const sel of FALLBACK_SELECTORS) if (await tryLoc(f.locator(sel))) return true;
  }
  return false;
}

/* ============= Human-like scroll =================== */
function normalizeMsRange(range, fallback) {
  const fallbackMin = fallback?.min || 10000;
  const fallbackMax = fallback?.max || fallbackMin;
  const min = Number.isFinite(range?.min) && range.min > 0 ? range.min : fallbackMin;
  const max = Number.isFinite(range?.max) && range.max >= min ? range.max : Math.max(min, fallbackMax);
  return { min, max };
}

async function humanScroll(page, range = HUMAN_SCROLL_MS) {
  try {
    const vp = page.viewportSize() || { width: 1366, height: 800 };
    const scrollRange = normalizeMsRange(range, HUMAN_SCROLL_MS);
    const TARGET_MS = rint(scrollRange.min, scrollRange.max);
    const startedAt = Date.now();
    const elapsed = () => Date.now() - startedAt;
    const remaining = () => TARGET_MS - elapsed();
    const readingMouseMove = async () => {
      try {
        const x = rint(Math.floor(vp.width * 0.15), Math.floor(vp.width * 0.82));
        const y = rint(Math.floor(vp.height * 0.20), Math.floor(vp.height * 0.78));
        await page.mouse.move(x, y, { steps: rint(8, 22) });
      } catch { }
    };
    const scrollBy = async (dy) => {
      await page.evaluate(d => window.scrollBy({ top: d, behavior: 'smooth' }), dy);
    };
    const nap = async (min, max) => {
      const want = rint(min, max);
      await sleep(Math.max(120, Math.min(want, remaining() + 400)));
    };
    await nap(500, 1300);
    await readingMouseMove();
    const initSteps = rint(2, 3);
    for (let i = 0; i < initSteps && remaining() > 1500; i++) {
      await scrollBy(rint(35, 95));
      await nap(380, 850);
      if (Math.random() < 0.45) await readingMouseMove();
    }
    while (remaining() > 1200) {
      const steps = rint(2, 4);
      for (let j = 0; j < steps && remaining() > 900; j++) {
        await scrollBy(rint(75, 190));
        await nap(100, 300);
      }
      const r = Math.random();
      if (r < 0.28) {
        await nap(1300, 3200);
        await readingMouseMove();
        await nap(350, 900);
      } else if (r < 0.62) {
        await nap(500, 1300);
        if (Math.random() < 0.5) await readingMouseMove();
      } else {
        await nap(180, 480);
      }
      if (remaining() > 2500 && Math.random() < 0.35) {
        await scrollBy(-rint(25, 90));
        await nap(300, 700);
        await readingMouseMove();
        await nap(220, 520);
        await scrollBy(rint(20, 60));
        await nap(180, 450);
      }
      if (remaining() > 4000 && Math.random() < 0.20) {
        await nap(350, 750);
        await scrollBy(-rint(120, 300));
        await nap(700, 1600);
        await readingMouseMove();
        await nap(300, 700);
      }
    }
    if (Math.random() < 0.6) {
      await scrollBy(rint(12, 40));
      await nap(180, 450);
      await scrollBy(-rint(5, 18));
      await nap(200, 450);
    }
    await readingMouseMove();
  } catch {
    // scroll failure is non-critical
  }
}

// *** REDUCED initial sleep from 1000 to 500 ***
async function humanizedStopGoScroll(page, range = HUMAN_SCROLL_MS) {
  try {
    const vp = page.viewportSize() || DESKTOP_VIEWPORT;
    const scrollRange = normalizeMsRange(range, HUMAN_SCROLL_MS);
    const pace = Math.max(0.75, Math.min(1.5, scrollRange.max / 15000));
    const nap = async (min, max) => sleep(Math.round(rint(min, max) * pace));

    const moveMouseReading = async () => {
      try {
        await page.mouse.move(
          rint(Math.floor(vp.width * 0.15), Math.floor(vp.width * 0.84)),
          rint(Math.floor(vp.height * 0.18), Math.floor(vp.height * 0.80)),
          { steps: rint(10, 24) }
        );
      } catch { }
    };

    const metrics = async () => page.evaluate(() => {
      const doc = document.documentElement;
      const body = document.body || doc;
      const scrollHeight = Math.max(body.scrollHeight, doc.scrollHeight, body.offsetHeight, doc.offsetHeight);
      return {
        y: window.scrollY || doc.scrollTop || 0,
        maxY: Math.max(0, scrollHeight - window.innerHeight),
        viewportH: window.innerHeight
      };
    }).catch(() => ({ y: 0, maxY: 0, viewportH: vp.height }));

    const slowScrollTo = async (targetY, minSteps, maxSteps) => {
      const start = await metrics();
      const boundedTarget = Math.max(0, Math.min(targetY, start.maxY));
      const distance = boundedTarget - start.y;
      if (Math.abs(distance) < 25) return;
      const calculatedSteps = Math.ceil(Math.abs(distance) / Math.max(150, start.viewportH * 0.28));
      const steps = Math.max(minSteps, Math.min(maxSteps, calculatedSteps));
      for (let i = 1; i <= steps; i++) {
        const progress = i / steps;
        const nextY = Math.round(start.y + distance * progress);
        await page.evaluate(y => window.scrollTo({ top: y, behavior: 'smooth' }), nextY);
        await nap(300, 620);
        if (Math.random() < 0.4) await moveMouseReading();
      }
      await page.evaluate(y => window.scrollTo({ top: y, behavior: 'auto' }), boundedTarget);
    };

    // *** REDUCED from 1000 to 500 ***
    await sleep(500);
    await moveMouseReading();

    const first = await metrics();
    if (first.maxY <= 0) return;
    const bottomGap = Math.min(first.viewportH * 0.65, first.maxY * 0.25);
    const partialDownY = Math.min(first.maxY - bottomGap, Math.max(first.viewportH * 0.9, first.maxY * 0.42));
    await slowScrollTo(partialDownY, 5, 12);
    await nap(2000, 3000);

    const partial = await metrics();
    const firstUpY = Math.max(0, partial.y - Math.max(partial.viewportH * 0.55, partial.y * 0.35));
    await slowScrollTo(firstUpY, 3, 7);
    await nap(1800, 2800);

    const beforeBottom = await metrics();
    await slowScrollTo(beforeBottom.maxY, 8, 18);
    await nap(2000, 3000);

    const atBottom = await metrics();
    const finalUpY = Math.max(0, atBottom.y - Math.max(atBottom.viewportH * 0.75, atBottom.maxY * 0.12));
    await slowScrollTo(finalUpY, 4, 9);
    await nap(1500, 2500);
    await moveMouseReading();
  } catch {
    // scroll failure is non-critical
  }
}

/* -------------- UA override + Client Hints --------- */
function extractChromeVersion(ua) {
  const m = ua.match(/Chrome\/(\d+)(?:\.(\d+)\.(\d+)\.(\d+))?/i);
  if (!m) return { major: '120', full: '120.0.0.0' };
  const major = m[1];
  const full = m[2] ? `${m[1]}.${m[2]}.${m[3]}.${m[4]}` : `${m[1]}.0.0.0`;
  return { major, full };
}
function detectPlatformFromUA(ua) {
  if (/Android/i.test(ua)) return { platform: 'Android', mobile: true, model: (ua.match(/\b([A-Z0-9-]{3,}) Build\//) || [])[1] || '', platformVersion: (ua.match(/Android\s([0-9.]+)/i) || [])[1] || '10' };
  if (/Windows NT/i.test(ua)) return { platform: 'Windows', mobile: false, model: '', platformVersion: (ua.match(/Windows NT\s([0-9.]+)/i) || [])[1] || '10.0' };
  if (/Macintosh/i.test(ua)) return { platform: 'macOS', mobile: false, model: '', platformVersion: '14.0.0' };
  if (/Linux x86_64|X11; Linux/i.test(ua)) return { platform: 'Linux', mobile: false, model: '', platformVersion: '' };
  return { platform: '', mobile: /Mobile/i.test(ua), model: '', platformVersion: '' };
}
function buildUAClientHints(ua) {
  const { major, full } = extractChromeVersion(ua);
  const { platform, mobile, model, platformVersion } = detectPlatformFromUA(ua);
  const brands = [
    { brand: "Chromium", version: major },
    { brand: "Google Chrome", version: major },
    { brand: "Not.A/Brand", version: "24" }
  ];
  const fullVersionList = [
    { brand: "Chromium", version: full },
    { brand: "Google Chrome", version: full },
    { brand: "Not.A/Brand", version: "24.0.0.0" }
  ];
  return { platform: platform || "", platformVersion: platformVersion || "", architecture: "", model: model || "", mobile: !!mobile, brands, fullVersionList, fullVersion: full };
}
async function overrideUA(page, ua) {
  if (!ua) return;
  try {
    const client = await page.context().newCDPSession(page);
    const ch = buildUAClientHints(ua);
    await client.send('Emulation.setUserAgentOverride', {
      userAgent: ua, platform: ch.platform || '', userAgentMetadata: ch
    });
  } catch (e) { console.warn('[UA override] Failed:', e.message); }
}

async function fitPageToViewport(page, viewport) {
  if (!page || !viewport) return;
  try {
    const client = await page.context().newCDPSession(page);
    await client.send('Emulation.clearDeviceMetricsOverride').catch(() => { });
    const { windowId } = await client.send('Browser.getWindowForTarget').catch(() => ({}));
    if (windowId) {
      await client.send('Browser.setWindowBounds', {
        windowId,
        bounds: { windowState: 'maximized' }
      }).catch(() => { });
    }
  } catch { }
  try {
    const size = await page.evaluate(() => ({ width: window.innerWidth, height: window.innerHeight })).catch(() => null);
    if (size?.width && size?.height) {
      console.log(`[Viewport] Browser content area ${size.width}x${size.height}`);
    }
  } catch { }
}

/* ---------------- IP log --------------------------- */
async function logPublicIP(context, label = 'IP') {
  try {
    const p = await context.newPage();
    if (FORCE_DESKTOP) { try { await p.setViewportSize(DESKTOP_VIEWPORT); } catch { } }
    await p.goto('https://api.ipify.org?format=json', { waitUntil: 'load', timeout: 20000 });
    console.log(`[IP ${label}] ${await p.textContent('body')}`);
    await p.close();
  } catch (e) { console.warn('IP check failed:', e.message); }
}

/* -------------- AdsPower group helpers ------------- */
function readGroupIdCache() { try { if (fs.existsSync(GROUP_ID_CACHE)) return fs.readFileSync(GROUP_ID_CACHE, 'utf8').trim(); } catch { } return ''; }
function writeGroupIdCache(id) { try { fs.writeFileSync(GROUP_ID_CACHE, String(id), 'utf8'); } catch { } }

async function listGroups() {
  const r = await httpJson('GET', `${ADSP_BASE}/api/v1/group/list?page=1&page_size=9999`);
  const list = (r.data?.list) || r.data || [];
  return list.map(g => ({ id: String(g.group_id || g.id), name: g.group_name || g.name || '' }));
}
async function ensureGroupId() {
  if (ADSP_GROUP_ID_ENV) return ADSP_GROUP_ID_ENV;
  const cached = readGroupIdCache();
  if (cached) return cached;
  try {
    const list = await listGroups();
    const found = list.find(g => g.name.toLowerCase() === ADSP_GROUP_NAME.toLowerCase());
    if (found) { writeGroupIdCache(found.id); return found.id; }
    if (list.length) { writeGroupIdCache(list[0].id); return list[0].id; }
  } catch (e) { console.warn('[AdsPower] group list warn:', e.message); }
  try {
    const c = await httpJson('POST', `${ADSP_BASE}/api/v1/group/create`, { group_name: ADSP_GROUP_NAME });
    if (c.code === 0 && (c.data?.group_id || c.data?.id)) {
      const gid = String(c.data.group_id || c.data.id);
      console.log('[AdsPower] Created group:', ADSP_GROUP_NAME, '→', gid);
      writeGroupIdCache(gid); return gid;
    }
    if ((c.msg || '').toLowerCase().includes('repeat') || (c.msg || '').toLowerCase().includes('exist')) {
      const list = await listGroups();
      const found = list.find(g => g.name.toLowerCase() === ADSP_GROUP_NAME.toLowerCase());
      if (found) { writeGroupIdCache(found.id); return found.id; }
    }
    throw new Error(c.msg || JSON.stringify(c));
  } catch (e) {
    throw new Error('Failed to acquire group_id (set ADSP_GROUP_ID or ADSP_GROUP_NAME). Cause: ' + e.message);
  }
}

/* -------- Build AdsPower proxy config for CREATE ---- */
function buildCreateProxyConfig(pObj) {
  if (!pObj) return { proxy_soft: 'no_proxy' };
  const u = new URL(pObj.server);
  const proxy_type = (/^socks/i.test(u.protocol)) ? 'socks5' : (u.protocol.includes('https') ? 'https' : 'http');
  return {
    proxy_soft: 'other',
    proxy_type,
    proxy_host: u.hostname,
    proxy_port: String(u.port || ''),
    proxy_user: pObj.username || '',
    proxy_password: pObj.password || ''
  };
}

/* -------- Rate limit helpers -------- */
function isRateLimit(msg = '') { const m = (msg || '').toLowerCase(); return m.includes('too many request') || m.includes('per second'); }
async function pollProfileIdByName(uniqueName, tries = 10) {
  for (let i = 0; i < tries; i++) {
    const r = await httpJson('GET', `${ADSP_BASE}/api/v1/user/list?page=1&page_size=500`);
    const list = (r.data?.list) || r.data || [];
    const hit = list.find(u => (u.name || '') === uniqueName);
    if (hit?.user_id) return String(hit.user_id);
    await sleep(300 + i * 200);
  }
  return null;
}

/* -------------- AdsPower profile ops --------------- */
function pickDesktopUA(uas, os) {
  const desktopList = (uas || []).filter(u =>
    /Windows NT|X11; Linux|Linux x86_64|Win64; x64/i.test(u) && !/Android|Mobile|iPhone|iPad/i.test(u)
  );
  if (desktopList.length) return choose(desktopList);
  return os === 'linux'
    ? 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
    : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36';
}
function pickMobileUA(uas) {
  const mobileList = (uas || []).filter(u => /Android|iPhone|iPad/i.test(u));
  if (mobileList.length) return choose(mobileList);
  return 'Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Mobile Safari/537.36';
}

async function adspowerCreateProfile(name, pObj, ua, mode = 'mobile') {
  const group_id = await ensureGroupId();
  let fingerprint_config;
  if (mode === 'desktop') {
    const osKey = ADSP_OS === 'linux' ? 'linux' : 'win';
    const navPlatform = ADSP_OS === 'linux' ? 'Linux x86_64' : 'Win32';
    fingerprint_config = {
      browser_type: "chrome", new_fingerprint: true,
      os: osKey, os_type: osKey, osType: osKey,
      platform: "desktop", device_type: "desktop", deviceType: "desktop",
      ua, navigator_platform: navPlatform
    };
  } else {
    fingerprint_config = {
      browser_type: "chrome", new_fingerprint: true,
      os: "android", os_type: "android", osType: "android",
      platform: "mobile", device_type: "mobile", deviceType: "mobile",
      ua, navigator_platform: "Linux armv8l"
    };
  }
  const body = {
    name, group_id,
    user_proxy_config: buildCreateProxyConfig(pObj),
    fingerprint_config
  };
  let backoff = 800;
  for (let attempt = 1; attempt <= 8; attempt++) {
    const r = await httpJson('POST', ADSP_BASE + '/api/v1/user/create', body);
    if (r.code === 0 && r.data?.user_id) {
      console.log('[AdsPower] Created profile →', r.data.user_id, `UA: ${ua.substring(0, 80)}...`);
      return String(r.data.user_id);
    }
    if (isRateLimit(r.msg)) {
      const wait = backoff + jitter(200, 600);
      console.log(`[AdsPower] Create rate-limited (attempt ${attempt}); wait ${wait}ms`);
      await sleep(wait); backoff = Math.min(backoff * 2, 8000); continue;
    }
    if ((r.msg || '').toLowerCase().includes('success')) {
      const pid = await pollProfileIdByName(name, 12);
      if (pid) { console.log('[AdsPower] Found created profile id via list →', pid); return pid; }
    }
    throw new Error('Create profile failed: ' + (r.msg || JSON.stringify(r)));
  }
  const pid = await pollProfileIdByName(name, 12);
  if (pid) { console.log('[AdsPower] Found created profile id via list (final) →', pid); return pid; }
  throw new Error('Create profile failed after retries.');
}

async function adspowerDeleteProfile(id) {
  try {
    const r = await httpJson('POST', ADSP_BASE + '/api/v1/user/delete', { user_ids: [id] });
    if (r.code !== 0) throw new Error(r.msg || JSON.stringify(r));
    console.log('[AdsPower] Deleted profile:', id);
  } catch (e) { console.warn('[AdsPower] Delete warn:', e.message); }
}
async function adspowerStop(id) { try { await httpJson('GET', ADSP_BASE + '/api/v1/browser/stop?user_id=' + encodeURIComponent(id)); } catch { } }
async function adspowerStartWithRetry(userId, { retries = 40, delayMs = 10000 } = {}) {
  const url = ADSP_BASE + '/api/v1/browser/start?user_id=' + encodeURIComponent(userId) + '&headless=' + (ADSP_HEADFUL ? 0 : 1);
  for (let i = 0; i < retries; i++) {
    const resp = await httpJson('GET', url);
    if (resp.code === 0 && resp.data?.ws?.puppeteer) return resp.data;
    const msg = (resp.msg || '').toLowerCase();
    if (msg.includes('updating') || msg.includes('downloading') || msg.includes('waiting') || msg.includes('installing')) {
      console.log(`[AdsPower] ${resp.msg} — retry ${i + 1}/${retries} in ${Math.round(delayMs / 1000)}s...`);
      await sleep(delayMs); continue;
    }
    if (msg.includes('already running')) { try { await adspowerStop(userId); } catch { } await sleep(1500); continue; }
    if (isRateLimit(msg)) { const wait = 1200 + jitter(200, 600); console.log(`[AdsPower] Start rate-limited; wait ${wait}ms`); await sleep(wait); continue; }
    throw new Error('AdsPower start failed: ' + JSON.stringify(resp));
  }
  throw new Error('AdsPower did not become ready in time.');
}
async function connectAdsPower(profileId) {
  await adspowerStop(profileId);
  const data = await adspowerStartWithRetry(profileId);
  const ws = data?.ws?.puppeteer || data?.ws?.playwright || data?.ws?.cdp || data?.ws?.wsUrl;
  if (!ws) throw new Error('No WebSocket endpoint from AdsPower response');
  console.log('[AdsPower] Connect over CDP:', ws);
  const browser = await chromium.connectOverCDP(ws);
  return { browser, stop: () => adspowerStop(profileId) };
}

/* -------------- Navigation helpers ---------------- */
// *** INCREASED to 180 seconds, added 429 detection ***
async function safeGoto(page, url) {
  try {
    const response = await page.goto(url, { waitUntil: 'load', timeout: 180000 });
    const status = response ? response.status() : 0;
    if (status === 429) {
      throw new Error('HTTP 429 Too Many Requests – proxy rate limited');
    }
    return true;
  } catch (e) {
    const m = String(e.message || '');
    if (m.includes('ERR_ABORTED')) { console.warn('[goto] ERR_ABORTED — treating as loaded.'); return true; }
    throw e;
  }
}

/* ======= Group runner — ONE group, ALL urls in parallel ======= */
function normalizeSiteUrl(url) {
  const raw = (url || '').trim();
  if (!raw) return '';
  return /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
}

function firstGroupUrl(groups, groupOrder) {
  for (const groupName of groupOrder || []) {
    const url = (groups[groupName] || []).find(notEmpty);
    if (url) return url;
  }
  return '';
}

function getSiteScrollRange() {
  const min = Number.isFinite(SITE_TEST_SCROLL_MS.min) && SITE_TEST_SCROLL_MS.min > 0 ? SITE_TEST_SCROLL_MS.min : 30000;
  const max = Number.isFinite(SITE_TEST_SCROLL_MS.max) && SITE_TEST_SCROLL_MS.max >= min ? SITE_TEST_SCROLL_MS.max : Math.max(min, 40000);
  return { min, max };
}

async function scrollWholePageForDuration(page, label, range = getSiteScrollRange()) {
  const targetMs = rint(range.min, range.max);
  console.log(`[${label}] Scrolling page for ${Math.round(targetMs / 1000)}s`);
  const startedAt = Date.now();
  let direction = 1;
  try { await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'auto' })); } catch { }
  while (Date.now() - startedAt < targetMs) {
    try {
      const metrics = await page.evaluate(() => {
        const doc = document.documentElement;
        const body = document.body || doc;
        const scrollHeight = Math.max(body.scrollHeight, doc.scrollHeight, body.offsetHeight, doc.offsetHeight);
        return {
          y: window.scrollY || doc.scrollTop || 0,
          maxY: Math.max(0, scrollHeight - window.innerHeight),
          width: window.innerWidth,
          height: window.innerHeight
        };
      });
      if (metrics.y >= metrics.maxY - 20) direction = -1;
      if (metrics.y <= 20) direction = 1;
      try {
        await page.mouse.move(
          rint(Math.floor(metrics.width * 0.15), Math.floor(metrics.width * 0.85)),
          rint(Math.floor(metrics.height * 0.18), Math.floor(metrics.height * 0.82)),
          { steps: rint(6, 18) }
        );
      } catch { }
      const distance = direction > 0 ? rint(180, 520) : -rint(90, 300);
      await page.evaluate(d => window.scrollBy({ top: d, behavior: 'smooth' }), distance);
      await sleep(rint(450, 1400));
    } catch {
      await sleep(500);
    }
  }
}

async function clickRandomInternalLink(page, siteUrl, label) {
  const links = await page.locator('a[href]').evaluateAll((anchors, base) => {
    const baseUrl = new URL(base);
    const current = window.location.href.replace(/#.*$/, '');
    return anchors.map((a, index) => {
      try {
        const url = new URL(a.href, window.location.href);
        const rect = a.getBoundingClientRect();
        return {
          index,
          href: url.href,
          text: (a.innerText || a.textContent || '').trim().slice(0, 80),
          visible: rect.width > 0 && rect.height > 0
        };
      } catch {
        return null;
      }
    }).filter(item =>
      item &&
      item.visible &&
      item.href.replace(/#.*$/, '') !== current &&
      new URL(item.href).origin === baseUrl.origin &&
      !/^(javascript:|mailto:|tel:)/i.test(item.href)
    );
  }, siteUrl).catch(() => []);

  if (!links.length) {
    console.log(`[${label}] No internal links found to click`);
    return false;
  }
  const chosen = choose(links);
  console.log(`[${label}] Random internal link: ${chosen.text || chosen.href}`);
  const link = page.locator('a[href]').nth(chosen.index);
  try {
    await link.scrollIntoViewIfNeeded({ timeout: 5000 });
    await sleep(rint(200, 600));
    const box = await link.boundingBox().catch(() => null);
    const hoverPoint = box ? {
      x: Math.round(box.x + box.width / 2),
      y: Math.round(box.y + Math.min(box.height * 0.55, box.height - 4))
    } : null;
    if (hoverPoint) {
      await page.mouse.move(hoverPoint.x, hoverPoint.y, { steps: rint(10, 24) }).catch(() => { });
      await sleep(rint(150, 400));
    }
    const clicked = await clickSmart(link, page);
    if (!clicked) throw new Error('clickSmart returned false');
    if (hoverPoint) {
      await page.mouse.move(hoverPoint.x, hoverPoint.y, { steps: 1 }).catch(() => { });
      await sleep(rint(100, 300));
    }
    try { await page.waitForLoadState('load', { timeout: 30000 }); } catch { }
    return true;
  } catch (e) {
    console.warn(`[${label}] Link click failed, opening href directly: ${e.message}`);
    await safeGoto(page, chosen.href);
    return true;
  }
}

async function runSiteTestInSameProfile(browser, siteUrl, fixedUA, viewport = DESKTOP_VIEWPORT) {
  const context = browser.contexts()[0];
  const page = await context.newPage();
  const label = 'SiteTest';
  try {
    await fitPageToViewport(page, viewport);
    page.on('dialog', d => d.accept().catch(() => { }));
    if (fixedUA) await overrideUA(page, fixedUA);
    await safeGoto(page, siteUrl);
    await forceRepaint(page);
    await sleep(rint(2500, 5000));
    await humanizedStopGoScroll(page, getSiteScrollRange());
    const clicked = await clickRandomInternalLink(page, siteUrl, label);
    if (clicked) {
      await forceRepaint(page);
      await sleep(rint(2500, 5000));
      await humanizedStopGoScroll(page, getSiteScrollRange());
    }
  } finally {
    try { await page.close(); } catch { }
    try { await context.clearCookies(); } catch { }
  }
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

async function typeIntoMockSearchVisibly(page, selector, query, label) {
  const searchBox = page.locator(selector);
  await searchBox.waitFor({ state: 'visible', timeout: 10000 });
  await page.bringToFront().catch(() => { });
  await searchBox.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
  await searchBox.click({ timeout: 5000, delay: 120 });
  await page.evaluate(({ sel, label }) => {
    const input = document.querySelector(sel);
    const status = document.getElementById('mockStatus');
    if (!input) return;
    input.focus();
    input.value = '';
    if (status) status.textContent = `${label}: ready to type`;
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }, { sel: selector, label });
  await sleep(rint(600, 1000));

  let current = '';
  for (const ch of query) {
    current += ch;
    await page.evaluate(({ current }) => {
      const status = document.getElementById('mockStatus');
      if (status) status.textContent = `Typing: ${current}`;
    }, { current });

    await searchBox.click({ timeout: 3000 }).catch(() => { });
    await page.keyboard.type(ch, { delay: rint(40, 90) }).catch(() => { });
    await sleep(rint(80, 180));

    const seen = await searchBox.inputValue().catch(() => '');
    if (seen !== current) {
      await searchBox.fill(current);
      await page.evaluate(({ sel, ch, current }) => {
        const input = document.querySelector(sel);
        const status = document.getElementById('mockStatus');
        if (!input) return;
        input.focus();
        input.value = current;
        if (status) status.textContent = `Typing: ${current}`;
        input.dispatchEvent(new KeyboardEvent('keydown', { key: ch, bubbles: true }));
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keyup', { key: ch, bubbles: true }));
      }, { sel: selector, ch, current });
    }
  }

  const typedValue = await searchBox.inputValue().catch(() => '');
  console.log(`[${label}] Search box value after typing: ${typedValue}`);
  if (typedValue !== query) {
    throw new Error(`Search typing failed: expected "${query}", got "${typedValue}"`);
  }
  await page.evaluate(() => {
    const status = document.getElementById('mockStatus');
    if (status) status.textContent = 'Typing complete. Waiting before search...';
  });
}

async function runMockSearchTestInSameProfile(browser, siteUrl, fixedUA, viewport = DESKTOP_VIEWPORT) {
  const context = browser.contexts()[0];
  const page = await context.newPage();
  const label = 'MockSearch';
  const exactName = MOCK_SEARCH_EXACT_NAME;
  const query = MOCK_SEARCH_QUERY;
  const safeName = escapeHtml(exactName);
  const safeUrl = escapeHtml(siteUrl);

  try {
    await fitPageToViewport(page, viewport);
    page.on('dialog', d => d.accept().catch(() => { }));
    if (fixedUA) await overrideUA(page, fixedUA);

    console.log(`[${label}] Loading local mock search page`);
    await page.goto('about:blank', { waitUntil: 'domcontentloaded', timeout: 10000 }).catch(() => { });
    const mockSearchHtml = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Training Search</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 42px auto; max-width: 760px; color: #202124; }
    form { display: flex; gap: 10px; margin-bottom: 28px; }
    input { flex: 1; font-size: 18px; padding: 12px 14px; border: 1px solid #dadce0; border-radius: 24px; }
    button { font-size: 15px; padding: 0 18px; border: 1px solid #dadce0; border-radius: 4px; background: #f8fafd; }
    #mockStatus { color: #5f6368; font-size: 14px; margin: -12px 0 22px; }
    .result { display: none; margin: 22px 0; }
    .result.show { display: block; }
    .result a { color: #1a0dab; font-size: 20px; text-decoration: none; }
    .url { color: #188038; font-size: 14px; margin-bottom: 4px; }
    .snippet { color: #4d5156; line-height: 1.45; }
  </style>
</head>
<body>
  <form id="searchForm">
    <input id="q" name="q" autocomplete="off" autofocus placeholder="Search training web" />
    <button type="submit">Search</button>
  </form>
  <div id="mockStatus">Preparing training search...</div>
  <main id="results">
    <div class="result" data-name="${safeName}">
      <div class="url">${safeUrl}</div>
      <a class="result-link" href="${safeUrl}">${safeName}</a>
      <div class="snippet">Exact training result for the configured site.</div>
    </div>
    <div class="result" data-name="example site ideas">
      <div class="url">https://example.test/ideas</div>
      <a href="https://example.test/ideas">example site ideas</a>
      <div class="snippet">Decoy result used to prove exact-name matching.</div>
    </div>
  </main>
  <script>
    const form = document.getElementById('searchForm');
    const input = document.getElementById('q');
    form.addEventListener('submit', event => {
      event.preventDefault();
      const q = input.value.trim().toLowerCase();
      document.getElementById('mockStatus').textContent = 'Showing results for: ' + input.value.trim();
      document.querySelectorAll('.result').forEach(result => {
        const name = result.dataset.name.toLowerCase();
        result.classList.toggle('show', name.includes(q) || q.includes(name));
      });
    });
  </script>
</body>
</html>`;
    await page.evaluate(html => {
      document.open();
      document.write(html);
      document.close();
    }, mockSearchHtml);
    await page.locator('#q').waitFor({ state: 'visible', timeout: 10000 });
    console.log(`[${label}] Local mock search page ready`);

    console.log(`[${label}] Typing search query: ${query}`);
    await typeIntoMockSearchVisibly(page, '#q', query, label);
    await sleep(rint(800, 1400));
    console.log(`[${label}] Clicking local Search button`);
    await page.locator('#searchForm button[type="submit"]').click({ timeout: 5000 });
    await page.waitForFunction(() => document.querySelectorAll('.result.show').length > 0, null, { timeout: 5000 })
      .catch(async () => {
        await page.evaluate(() => {
          const form = document.getElementById('searchForm');
          if (form) form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
        }).catch(() => { });
      });

    const exactResult = page.locator('.result.show .result-link', { hasText: new RegExp(`^\\s*${exactName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*$`, 'i') }).first();
    await exactResult.waitFor({ state: 'visible', timeout: 5000 });
    console.log(`[${label}] Exact result found: ${exactName}`);

    console.log(`[${label}] Clicking exact result and navigating to site`);
    const resultHref = await exactResult.getAttribute('href');
    await exactResult.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
    await sleep(rint(400, 900));
    await exactResult.click({ timeout: 6000, noWaitAfter: true });
    await page.waitForURL(u => u.href !== 'about:blank', { timeout: 6000 }).catch(() => { });
    if (resultHref && page.url() === 'about:blank') {
      console.warn(`[${label}] Result click did not navigate quickly; opening href directly`);
      await safeGoto(page, resultHref);
    } else {
      await page.waitForLoadState('load', { timeout: 45000 }).catch(() => { });
    }
    console.log(`[${label}] Current URL after result click: ${page.url()}`);

    await forceRepaint(page);
    await sleep(rint(2500, 5000));
    await humanizedStopGoScroll(page, getSiteScrollRange());

    const clicked = await clickRandomInternalLink(page, siteUrl, label);
    if (clicked) {
      await forceRepaint(page);
      await sleep(rint(2500, 5000));
      await humanizedStopGoScroll(page, getSiteScrollRange());
    }
    console.log(`[${label}] Flow complete`);
  } catch (e) {
    console.error(`[${label}] Failed: ${e.message}`);
    try {
      await page.screenshot({ path: path.join(__dirname, 'mock-search-debug.png'), fullPage: true });
      console.error(`[${label}] Saved debug screenshot: mock-search-debug.png`);
    } catch { }
    throw e;
  } finally {
    if (Number.isFinite(MOCK_SEARCH_KEEP_OPEN_MS) && MOCK_SEARCH_KEEP_OPEN_MS > 0) {
      console.log(`[${label}] Keeping browser open for ${Math.round(MOCK_SEARCH_KEEP_OPEN_MS / 1000)}s before cleanup`);
      await sleep(MOCK_SEARCH_KEEP_OPEN_MS);
    }
    try { await page.close(); } catch { }
    try { await context.clearCookies(); } catch { }
  }
}

const pageUAStore = new WeakMap();

// Force GPU repaint — AdsPower CDP sessions can render blank/dark without this.
async function forceRepaint(page) {
  try {
    await page.evaluate(() => new Promise(resolve => {
      document.documentElement.style.zoom = '0.75';
      requestAnimationFrame(() => {
        document.documentElement.style.removeProperty('zoom');
        requestAnimationFrame(resolve);
      });
    })).catch(() => { });
    await page.evaluate(() => { window.scrollTo(0, 1); window.scrollTo(0, 0); }).catch(() => { });
    await page.screenshot({ type: 'jpeg', quality: 1 }).catch(() => { });
  } catch { }
}

/* ================== runPageFlow – Mobile‑optimised ================== */
async function runPageFlow(page, url, groupLabel, context, fixedUA, viewport) {
  if (fixedUA) { await overrideUA(page, fixedUA); }
  await fitPageToViewport(page, viewport);

  const startTime = Date.now();
  const minDuration = 30000; // 30 seconds minimum total time

  try {
    await forceRepaint(page);
    await sleep(1000);

    console.log(`[${groupLabel}] Skipping CTA clicks – only internal links`);

    // First settle and scroll
    const settle = rint(1000, 3000);
    console.log(`[${groupLabel}] Settling ${Math.round(settle / 1000)}s before first scroll`);
    await sleep(settle);

    console.log(`[${groupLabel}] First scroll (8–15s)`);
    try {
      await humanizedStopGoScroll(page);
    } catch (e) {
      console.warn(`[${groupLabel}] First scroll error: ${e.message}`);
    }
    await sleep(rint(2000, 4000));

    // --------------------------
    // CLICK INTERNAL LINK – on both desktop AND mobile
    // --------------------------
    console.log(`[${groupLabel}] Clicking random internal link (blog post)`);
    let internalClicked = false;
    try {
      internalClicked = await clickRandomInternalLink(page, url, groupLabel);
    } catch (e) {
      console.warn(`[${groupLabel}] Internal link click error: ${e.message}`);
    }

    if (internalClicked) {
      // We are now on the blog post
      await forceRepaint(page);
      console.log(`[${groupLabel}] On blog post – waiting 3-5s to simulate reading`);
      await sleep(rint(3000, 5000));

      // ----- GO BACK TO HOMEPAGE -----
      console.log(`[${groupLabel}] Returning to homepage: ${url}`);
      let returned = false;
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          await safeGoto(page, url);
          await forceRepaint(page);
          await sleep(rint(2000, 4000));
          returned = true;
          console.log(`[${groupLabel}] Successfully returned to homepage (attempt ${attempt})`);
          break;
        } catch (e) {
          console.warn(`[${groupLabel}] Return attempt ${attempt} failed: ${e.message}`);
          if (attempt < 3) await sleep(2000);
        }
      }
      if (!returned) {
        // Last resort: reload current page
        console.warn(`[${groupLabel}] Could not return to homepage – reloading current page`);
        try {
          await page.reload({ timeout: 30000 });
          await forceRepaint(page);
          await sleep(rint(2000, 4000));
        } catch (e) {
          console.warn(`[${groupLabel}] Reload failed: ${e.message}`);
        }
      }
    } else {
      console.log(`[${groupLabel}] No internal link found – staying on homepage`);
    }

    // --------------------------
    // Final scroll and dwell – ALWAYS executed
    // --------------------------
    console.log(`[${groupLabel}] Final scroll (12–16s) on current page: ${page.url()}`);
    try {
      await humanizedStopGoScroll(page, INTERNAL_LINK_SCROLL_MS);
    } catch (e) {
      console.warn(`[${groupLabel}] Final scroll error: ${e.message}`);
    }

    const dwell = rint(DWELL_MS.min, DWELL_MS.max);
    console.log(`[${groupLabel}] Final dwell ${Math.round(dwell / 1000)}s`);
    await sleep(dwell);

    // Ensure we don't finish too early
    const elapsed = Date.now() - startTime;
    if (elapsed < minDuration) {
      const extra = minDuration - elapsed;
      console.log(`[${groupLabel}] Adding extra ${Math.round(extra / 1000)}s to meet minimum duration`);
      await sleep(extra);
    }

    console.log(`[${groupLabel}] Finished – ended on: ${page.url()}`);

  } catch (e) {
    console.warn(`[${groupLabel}] Flow error: ${e.message}`);
    const elapsed = Date.now() - startTime;
    if (elapsed < minDuration) {
      const extra = minDuration - elapsed;
      console.log(`[${groupLabel}] Error – waiting extra ${Math.round(extra / 1000)}s before closing`);
      await sleep(extra);
    }
  }
}

/* ================== googleSearchAndLand – with reload fallback ================== */
async function googleSearchAndLand(browser, query, targetUrl, fixedUA, viewport) {
  const context = browser.contexts()[0];
  const page = await context.newPage();
  try {
    await fitPageToViewport(page, viewport);
    if (fixedUA) await overrideUA(page, fixedUA);
    page.on('dialog', d => d.accept().catch(() => { }));

    // 1. Go to Google
    await safeGoto(page, 'https://www.google.com/');
    await forceRepaint(page);
    await sleep(rint(1000, 3000));

    // 2. Accept cookies
    const consent = page.locator('button:has-text("Accept all"), button:has-text("I agree")').first();
    if (await consent.isVisible({ timeout: 3000 }).catch(() => false)) {
      await consent.click();
      await sleep(rint(500, 1500));
    }

    // 3. Type query
    const searchSelectors = [
      'textarea[name="q"]',
      'input[name="q"]',
      'input[aria-label="Search"]',
      'input.gLFyf',
      'input[role="combobox"]'
    ];
    let searchBox = null;
    for (const sel of searchSelectors) {
      const loc = page.locator(sel).first();
      if (await loc.isVisible({ timeout: 1000 }).catch(() => false)) {
        searchBox = loc;
        break;
      }
    }
    if (!searchBox) searchBox = page.locator('input[type="text"], textarea').first();
    await searchBox.waitFor({ state: 'visible', timeout: 30000 });
    await searchBox.scrollIntoViewIfNeeded();
    await sleep(200);
    await searchBox.focus();
    await sleep(300);

    for (const ch of query) {
      await page.keyboard.type(ch, { delay: rint(40, 90) });
      await sleep(rint(50, 120));
    }
    await sleep(rint(400, 800));
    await page.keyboard.press('Enter');

    // ---- CAPTCHA handling with retry ----
    let retries = 0;
    const maxRetries = 2;
    while (retries <= maxRetries) {
      const currentUrl = page.url();
      if (currentUrl.includes('/sorry/') || currentUrl.includes('recaptcha')) {
        console.log(`[GoogleSearch] CAPTCHA detected (attempt ${retries + 1}/${maxRetries + 1}) – waiting...`);
        let solved = false;
        for (let i = 0; i < 150; i++) {
          await sleep(2000);
          const url = page.url();
          if (!url.includes('/sorry/') && !url.includes('recaptcha')) {
            solved = true;
            break;
          }
        }
        if (!solved) {
          console.log('[GoogleSearch] Extension did not solve – using bot solver');
          await solveGoogleCaptcha(page);
        } else {
          console.log('[GoogleSearch] Extension solved CAPTCHA');
        }
      }

      // Wait for results with timeout
      console.log('[GoogleSearch] Waiting for results (timeout: 120s)...');
      try {
        await page.waitForSelector('#search, #rso, #res, .g', { timeout: 120000 });
        break; // success
      } catch (e) {
        console.warn(`[GoogleSearch] No results after 120s (attempt ${retries + 1}) – reloading...`);
        // ---- Clear extension session before reload ----
        await page.evaluate(() => {
          sessionStorage.removeItem('captchaSolverSession');
        }).catch(() => { });
        await page.reload({ timeout: 60000 });
        await forceRepaint(page);
        retries++;
      }
    }

    // Final check
    if (!(await page.locator('#search, #rso, #res, .g').first().isVisible({ timeout: 5000 }).catch(() => false))) {
      throw new Error('No search results after multiple retries');
    }

    await forceRepaint(page);
    await sleep(rint(2000, 5000));

    // Scroll results
    console.log('[GoogleSearch] Scrolling Google results page before clicking');
    await humanizedStopGoScroll(page, { min: 3000, max: 6000 });

    // Find and click result (same logic as before)
    const domain = targetUrl.replace(/^https?:\/\//, '').split('/')[0];
    const containerSelectors = ['#rso', '#res', '#search', '.g', 'div[role="main"]'];
    let links = [];
    for (const sel of containerSelectors) {
      const container = page.locator(sel).first();
      if (await container.isVisible({ timeout: 1000 }).catch(() => false)) {
        const containerLinks = container.locator('a[href]');
        const count = await containerLinks.count();
        if (count > 0) {
          for (let i = 0; i < count; i++) {
            const link = containerLinks.nth(i);
            const href = await link.getAttribute('href').catch(() => null);
            const text = await link.textContent().catch(() => '');
            links.push({ link, href, text, visible: true });
          }
          break;
        }
      }
    }

    if (links.length === 0) {
      console.log('[GoogleSearch] No links in containers – falling back to all page links');
      const allLinks = page.locator('a[href]');
      const count = await allLinks.count();
      for (let i = 0; i < count; i++) {
        const link = allLinks.nth(i);
        const href = await link.getAttribute('href').catch(() => null);
        const text = await link.textContent().catch(() => '');
        if (href && !href.startsWith('#') && !href.startsWith('javascript:') && !href.startsWith('/search?') && !href.includes('accounts.google.com')) {
          links.push({ link, href, text, visible: true });
        }
      }
    }

    console.log(`[GoogleSearch] Found ${links.length} candidate links`);

    let resultLink = null;
    let href = null;
    for (const item of links) {
      if (item.href && item.href.includes(domain) && !item.href.startsWith('#') && !item.href.startsWith('javascript:')) {
        const classes = await item.link.getAttribute('class').catch(() => '');
        if (classes && (classes.includes('ad') || classes.includes('sponsored'))) continue;
        resultLink = item.link;
        href = item.href;
        break;
      }
    }
    if (!resultLink) {
      for (const item of links) {
        const text = item.text.trim();
        if (text && text.length > 0 && !text.includes('Ads') && !text.includes('Sponsored') && !text.includes('People also ask')) {
          const classes = await item.link.getAttribute('class').catch(() => '');
          if (classes && (classes.includes('ad') || classes.includes('sponsored'))) continue;
          resultLink = item.link;
          href = item.href;
          break;
        }
      }
    }
    if (!resultLink) {
      for (const item of links) {
        if (item.href && item.href.startsWith('http') && !item.href.includes('google.com')) {
          const classes = await item.link.getAttribute('class').catch(() => '');
          if (classes && (classes.includes('ad') || classes.includes('sponsored'))) continue;
          resultLink = item.link;
          href = item.href;
          break;
        }
      }
    }

    if (!resultLink) {
      console.log('[GoogleSearch] Debug: page URL:', await page.url());
      const allHrefs = await page.evaluate(() => Array.from(document.querySelectorAll('a[href]'), a => a.href));
      console.log(allHrefs.slice(0, 20));
      throw new Error('No search results found – no usable links');
    }

    console.log(`[GoogleSearch] Clicking result → ${href || 'unknown'}`);
    await resultLink.scrollIntoViewIfNeeded();
    await sleep(rint(500, 1200));

    try {
      await Promise.all([
        page.waitForLoadState('load', { timeout: 180000 }).catch(() => { }),
        resultLink.click({ timeout: 6000 })
      ]);
    } catch (e) {
      if (href && href.startsWith('http')) {
        console.warn('[GoogleSearch] Click failed, navigating directly to href');
        await safeGoto(page, href);
      } else if (href && href.startsWith('/')) {
        const base = page.url().match(/^https?:\/\/[^\/]+/)[0];
        await safeGoto(page, base + href);
      } else {
        throw e;
      }
    }

    await forceRepaint(page);
    await sleep(rint(2000, 5000));
    console.log(`[GoogleSearch] Landed on: ${page.url()}`);
    return page;
  } catch (e) {
    await page.close().catch(() => { });
    throw e;
  }
}

async function runGoogleSearchAndFlow(browser, query, targetUrl, groupLabel, fixedUA, viewport) {
  const context = browser.contexts()[0];
  const page = await googleSearchAndLand(browser, query, targetUrl, fixedUA, viewport);
  try {
    await runPageFlow(page, targetUrl, groupLabel, context, fixedUA, viewport);
  } finally {
    try { await page.close(); } catch { }
    try { await context.clearCookies(); } catch { }
  }
}

/* ======= Group runner — ONE group, ALL urls in parallel ======= */
async function runGroupInSameProfile(browser, groupLabel, urls, fixedUA, viewport = DESKTOP_VIEWPORT) {
  const context = browser.contexts()[0];

  for (const p of context.pages()) {
    if (fixedUA) { await overrideUA(p, fixedUA); pageUAStore.set(p, fixedUA); }
    await fitPageToViewport(p, viewport);
    p.on('dialog', d => d.accept().catch(() => { }));
  }

  const onNewPage = async (newPage) => {
    let ua = fixedUA;
    const opener = newPage.opener();
    if (opener && pageUAStore.has(opener)) ua = pageUAStore.get(opener);
    if (ua) { await overrideUA(newPage, ua); pageUAStore.set(newPage, ua); }
    await fitPageToViewport(newPage, viewport);
    newPage.on('dialog', d => d.accept().catch(() => { }));
  };
  context.on('page', onNewPage);

  const runOne = async (url) => {
    await sleep(rint(PRENAV_PAUSE_MS.min, PRENAV_PAUSE_MS.max));
    const page = await context.newPage();
    await fitPageToViewport(page, viewport);
    if (fixedUA) { await overrideUA(page, fixedUA); pageUAStore.set(page, fixedUA); }

    try {
      const uaSeen = await page.evaluate(() => navigator.userAgent);
      console.log(`[${groupLabel}] UA: ${uaSeen.substring(0, 180)}${uaSeen.length > 180 ? '...' : ''}`);
    } catch { }

    let ok = false;
    for (let attempt = 0; attempt <= RETRY_NAVIGATION; attempt++) {
      try {
        const loaded = await safeGoto(page, url);
        if (loaded) {
          await runPageFlow(page, url, groupLabel, context, fixedUA, viewport);
        }
        ok = true; break;
      } catch (e) {
        console.warn(`[${groupLabel}] Error on ${url}: ${e.message}`);
        if (attempt < RETRY_NAVIGATION) await sleep(1000);
      }
    }
    try { await page.close(); } catch { }
    if (!ok) console.warn(`[${groupLabel}] Failed after retries: ${url}`);
  };

  try {
    const limit = (GROUP_MAX_TABS && GROUP_MAX_TABS > 0) ? GROUP_MAX_TABS : urls.length;
    for (let i = 0; i < urls.length; i += limit) {
      const batch = urls.slice(i, i + limit);
      await Promise.allSettled(batch.map(u => runOne(u)));
    }
  } finally {
    context.off('page', onNewPage);
    try { await context.clearCookies(); } catch { }
  }
}

/* ====================== Main ======================= */
(async () => {
  const { order: allGroupOrder, groups } = loadGroups(GROUPS_FILE);

  const selectedGroups = GROUPS_PER_PROFILE.length
    ? allGroupOrder.filter(g => GROUPS_PER_PROFILE.includes(g))
    : allGroupOrder;
  const groupsToRun = selectedGroups.length ? selectedGroups : allGroupOrder;
  const groupDefaultUrl = firstGroupUrl(groups, groupsToRun);
  const siteTestUrl = normalizeSiteUrl(SITE_TEST_URL || ADSCFG.siteTestUrl || groupDefaultUrl);
  if ((SITE_TEST_USE || MOCK_SEARCH_USE) && !siteTestUrl) {
    throw new Error('No target URL found. Add a URL to groups.json or set SITE_TEST_URL.');
  }
  console.log(`Groups in cycle order: ${groupsToRun.join(' → ')}`);

  if (SITE_TEST_USE || MOCK_SEARCH_USE) {
    const scrollRange = getSiteScrollRange();
    if (MOCK_SEARCH_USE) {
      console.log(`[MockSearch] Training search enabled: query="${MOCK_SEARCH_QUERY}", exact="${MOCK_SEARCH_EXACT_NAME}", target=${siteTestUrl}`);
    } else {
      console.log(`[SiteTest] Direct site test enabled: ${siteTestUrl}`);
    }
    console.log(`[SiteTest] Scroll range: ${Math.round(scrollRange.min / 1000)}-${Math.round(scrollRange.max / 1000)}s per page`);
  }

  const { proxies: proxiesX, uas: uasX, desktopUas: desktopUasX } = loadFromExcel(PROFILES_XLSX);
  const proxies = proxiesX.length ? proxiesX : loadLines(PROXIES_FILE);
  const uas = (uasX.length ? uasX : loadLines(UAS_FILE)).filter(notEmpty);
  const desktopUas = (desktopUasX.length ? desktopUasX : loadLines(DESKTOP_UAS_FILE)).filter(notEmpty);

  const parsedProxies = proxies.map(parseProxyLine).filter(Boolean);
  console.log(`Loaded ${parsedProxies.length} proxies, ${uas.length} mobile UA, ${desktopUas.length} desktop UA`);

  if (proxies.length > 0 && parsedProxies.length === 0) {
    console.error('ERROR: Proxy file has entries but none could be parsed. Fix the format before running.');
    console.error('Expected formats: user:pass:host:port  OR  user:pass@host:port  OR  host:port');
    process.exit(1);
  }

  if (IPQS_KEYS.length) {
    const exhaustedToday = loadExhaustedKeys();
    const before = IPQS_KEYS.length;
    IPQS_KEYS = IPQS_KEYS.filter(k => !exhaustedToday.has(k));
    const skipped = before - IPQS_KEYS.length;
    if (skipped > 0) {
      console.log(`[IPQS] ${skipped} key(s) already hit daily limit today — skipped. ${IPQS_KEYS.length} key(s) available.`);
    }
  }

  if (IPQS_KEYS.length) {
    console.log(`[IPQS] Proxy scoring enabled — ${IPQS_KEYS.length} key(s), strictness=${IPQS_STRICTNESS}, max score=${IPQS_MAX_SCORE}`);
  } else {
    console.log('[IPQS] No API key set — proxy scoring disabled, using proxies as-is');
  }

  let cycles = parseInt(await ask('How many profiles to run? '), 10);
  if (!Number.isFinite(cycles) || cycles <= 0) { cycles = 1; console.log('Using default: 1'); }

  const modeAnswer = (await ask('Run Mobile or Desktop? (m/d): ')).trim().toLowerCase();
  const runMode = (modeAnswer === 'd' || modeAnswer === 'desktop') ? 'desktop' : 'mobile';
  console.log(`[Mode] ${runMode === 'desktop' ? 'Desktop' : 'Mobile (Android/iOS)'} profiles selected`);

  const ipqsAnswer = IPQS_KEYS.length
    ? (await ask('Enable IPQS proxy check? (y/n): ')).trim().toLowerCase()
    : 'n';
  const useIPQS = ipqsAnswer === 'y' || ipqsAnswer === 'yes';
  if (IPQS_KEYS.length) console.log(useIPQS ? '[IPQS] Enabled — bad proxies will be skipped.' : '[IPQS] Disabled — all proxies used without scoring.');

  let proxyAbsIndex = 0;
  let groupIndex = 0;
  let completedProfiles = 0;
  let consecutiveBad = 0;

  while (completedProfiles < cycles) {
    if (parsedProxies.length > 0 && consecutiveBad >= parsedProxies.length) {
      console.error('[IPQS] All proxies produced bad exit IPs. Stopping.');
      break;
    }
    if (!parsedProxies.length) {
      console.error('No proxies available. Refusing to open browser on main IP. Stopping.');
      break;
    }

    const pObj = parsedProxies[proxyAbsIndex % parsedProxies.length];
    proxyAbsIndex++;

    if (useIPQS) {
      const exitIP = await getExitIPViaProxy(pObj);
      if (exitIP) {
        const score = await checkWithKeyRotation(exitIP);
        if (!IPQS_KEYS.length) {
          console.error('[IPQS] All API keys exhausted for today — stopping run. Restart tomorrow.');
          break;
        }
        if (score === null) {
          console.warn(`[IPQS] Could not score ${exitIP} (API error) — skipping proxy, no browser opened`);
          consecutiveBad++;
          continue;
        }
        if (score >= IPQS_MAX_SCORE) {
          console.log(`[IPQS] ${exitIP} score=${score} ✗ BAD — skip proxy, no browser opened`);
          consecutiveBad++;
          continue;
        }
        console.log(`[IPQS] ${exitIP} score=${score} ✓ GOOD`);
      } else {
        console.warn('[IPQS] All IP-check services failed for this proxy — skipping, no browser opened');
        consecutiveBad++;
        continue;
      }
    }
    consecutiveBad = 0;

    const gName = groupsToRun[groupIndex % groupsToRun.length];
    groupIndex++;
    completedProfiles++;

    const urls = (groups[gName] || []).filter(notEmpty);
    const profileUA = runMode === 'desktop'
      ? (choose(desktopUas) || pickDesktopUA(desktopUas.length ? desktopUas : uas, ADSP_OS))
      : (pickMobileUA(uas) || choose(uas));
    const profileViewport = runMode === 'desktop' ? DESKTOP_VIEWPORT : choose(MOBILE_SIZES);

    const display = { server: pObj.server };
    if (pObj.username) display.user = '***';
    console.log(`\n=== Profile ${completedProfiles}/${cycles} — Group ${gName} (${urls.length} URL${urls.length !== 1 ? 's' : ''})`);
    console.log(`[Proxy] ${JSON.stringify(display)} | Screen ${profileViewport.width}x${profileViewport.height}`);
    let profileId = null, browser = null, stopper = null;

    try {
      if (ADSP_USE) {
        const uniqueName = `Auto-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
        profileId = await adspowerCreateProfile(uniqueName, pObj, profileUA, runMode);
        const connected = await connectAdsPower(profileId);
        browser = connected.browser;
        stopper = connected.stop;
      } else {
        const launchOptions = { headless: !HEADFUL, args: [`--window-size=${DESKTOP_VIEWPORT.width},${DESKTOP_VIEWPORT.height}`] };
        launchOptions.proxy = { server: pObj.server };
        if (pObj.username) launchOptions.proxy.username = pObj.username;
        if (pObj.password) launchOptions.proxy.password = pObj.password;
        browser = await chromium.launch(launchOptions);
      }

      if (MOCK_SEARCH_USE) {
        const scrollRange = getSiteScrollRange();
        const mockSearchTimeoutMs = Math.max(PROFILE_TIMEOUT_MS, scrollRange.max * 2 + 80000);
        const timeoutGuard = sleep(mockSearchTimeoutMs).then(() => {
          throw new Error(`Profile timed out after ${Math.round(mockSearchTimeoutMs / 1000)}s - mock search test took too long`);
        });
        await Promise.race([
          runMockSearchTestInSameProfile(browser, siteTestUrl, profileUA, profileViewport),
          timeoutGuard,
        ]);
      } else if (SITE_TEST_USE) {
        const scrollRange = getSiteScrollRange();
        const siteTestTimeoutMs = Math.max(PROFILE_TIMEOUT_MS, scrollRange.max * 2 + 70000);
        const timeoutGuard = sleep(siteTestTimeoutMs).then(() => {
          throw new Error(`Profile timed out after ${Math.round(siteTestTimeoutMs / 1000)}s - site test took too long`);
        });
        await Promise.race([
          runSiteTestInSameProfile(browser, siteTestUrl, profileUA, profileViewport),
          timeoutGuard,
        ]);
      } else if (GOOGLE_SEARCH_USE) {
        const targetUrl = urls.length ? urls[0] : 'example.com';
        await runGoogleSearchAndFlow(browser, GOOGLE_SEARCH_QUERY, targetUrl, `Google→${gName}`, profileUA, profileViewport);
      } else if (urls.length === 0) {
        console.warn(`Group ${gName} has no URLs — skipping`);
      } else {
        const timeoutGuard = sleep(PROFILE_TIMEOUT_MS).then(() => {
          throw new Error(`Profile timed out after ${PROFILE_TIMEOUT_MS / 1000}s — proxy likely died`);
        });
        await Promise.race([
          runGroupInSameProfile(browser, `Group ${gName}`, urls, profileUA, profileViewport),
          timeoutGuard,
        ]);
      }

    } catch (err) {
      console.error(`Profile ${completedProfiles} error:`, err.message);
    } finally {
      if (browser) { try { await Promise.race([browser.close(), sleep(10000)]); } catch { } }
      if (stopper) { try { await stopper(); } catch { } }
      if (ADSP_USE && profileId) { await adspowerDeleteProfile(profileId); }
      await sleep(jitter(INTER_RUN_SLEEP_MS.min, INTER_RUN_SLEEP_MS.max));
    }
  }

  console.log('\nAll profiles finished. Exiting.');
  process.exit(0);
})().catch(e => { console.error(e); process.exit(1); });