'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const DATA_DIR = path.join(os.homedir(), '.sp-api');
const PROFILE_DIR = path.join(DATA_DIR, 'browser-profile');
const AUTH_FILE = path.join(DATA_DIR, 'auth.json');
const LOGIN_TIMEOUT_MS = 300_000;
const HEADLESS_PROBE_MS = 5_000;

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  setPrivateMode(dir, 0o700);
}

function setPrivateMode(target, mode) {
  if (process.platform === 'win32') return;
  fs.chmodSync(target, mode);
}

function readAuthFile(authFile = AUTH_FILE) {
  try {
    return JSON.parse(fs.readFileSync(authFile, 'utf8'));
  } catch {
    return null;
  }
}

function authStatus(authFile = AUTH_FILE) {
  const parsed = readAuthFile(authFile);
  return {
    authFile,
    exists: !!parsed,
    site: parsed?.SP_SITE || null,
    hasCookies: !!parsed?.SP_COOKIES,
    hasToken: !!parsed?.SP_TOKEN,
  };
}

function requirePlaywright() {
  try {
    return require('playwright');
  } catch {
    throw new Error('playwright is not installed. Run npm install in the sp-api repo.');
  }
}

function isLoginUrl(url) {
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host.includes('login.microsoftonline.com')
      || host.includes('login.microsoft.com')
      || host.includes('login.live.com');
  } catch {
    return false;
  }
}

function parseSiteInput(raw) {
  const cleaned = raw.replace(/^https?:\/\//, '').replace(/\/+$/, '');
  const slashIdx = cleaned.indexOf('/');
  if (slashIdx === -1) {
    return { tenantHost: cleaned, sitePath: '' };
  }
  return {
    tenantHost: cleaned.substring(0, slashIdx),
    sitePath: cleaned.substring(slashIdx),
  };
}

function buildCookieString(cookies, tenantHost) {
  const domainCookies = cookies.filter(c => {
    const domain = c.domain.replace(/^\./, '');
    return tenantHost.endsWith(domain) || domain.endsWith(tenantHost);
  });
  const authNames = new Set(['FedAuth', 'rtFa', 'SPOIDCRL', 'CcsAuth']);
  const authCookies = domainCookies.filter(c => authNames.has(c.name) || c.name.startsWith('FedAuth'));
  const chosen = authCookies.length > 0 ? authCookies : domainCookies;
  return chosen.map(c => `${c.name}=${c.value}`).join('; ');
}

async function authenticate(site, options = {}) {
  const { forceLogin = false, verbose = false, playwright = requirePlaywright(), authFile = AUTH_FILE, profileDir = PROFILE_DIR } = options;
  const { tenantHost, sitePath } = parseSiteInput(site);
  ensureDir(profileDir);

  const tenantUrl = `https://${tenantHost}`;
  const siteUrl = sitePath ? `${tenantUrl}${sitePath}` : tenantUrl;
  let headless = !forceLogin;
  const log = (msg) => { if (verbose) process.stderr.write(`[auth] ${msg}\n`); };

  log(`launching Edge (${headless ? 'headless probe' : 'visible'}, profile=${profileDir})`);
  let context = await playwright.chromium.launchPersistentContext(profileDir, {
    channel: 'msedge',
    headless,
    args: ['--disable-blink-features=AutomationControlled'],
    viewport: { width: 1280, height: 800 },
  });
  let page = context.pages()[0] || await context.newPage();

  const capturedTokens = { sp: null, spScopes: 0 };
  function classifyToken(token) {
    try {
      const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
      const aud = (payload.aud || '').toLowerCase();
      const scopes = (payload.scp || '').split(' ').length;
      if (aud.includes('.sharepoint.')) return { type: 'sp', scopes };
    } catch {}
    return null;
  }
  function installTokenInterceptor(targetPage) {
    targetPage.on('request', request => {
      const auth = request.headers().authorization;
      if (!auth?.startsWith('Bearer ')) return;
      const token = auth.substring(7);
      const info = classifyToken(token);
      if (info?.type === 'sp' && info.scopes > capturedTokens.spScopes) {
        capturedTokens.sp = token;
        capturedTokens.spScopes = info.scopes;
      }
    });
  }
  installTokenInterceptor(page);

  log(`loading ${siteUrl}`);
  await page.goto(siteUrl, { waitUntil: 'domcontentloaded' });
  if (headless && isLoginUrl(page.url())) {
    log('login redirect detected — relaunching visible Edge for sign-in');
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: HEADLESS_PROBE_MS });
    } catch {
      await context.close();
      process.stderr.write('Opening Edge for login. Complete sign-in in the browser window.\n');
      context = await playwright.chromium.launchPersistentContext(profileDir, {
        channel: 'msedge',
        headless: false,
        args: ['--disable-blink-features=AutomationControlled'],
        viewport: { width: 1280, height: 800 },
      });
      page = context.pages()[0] || await context.newPage();
      installTokenInterceptor(page);
      await page.goto(siteUrl, { waitUntil: 'domcontentloaded' });
      headless = false;
    }
  }

  if (!headless && isLoginUrl(page.url())) {
    log(`waiting up to ${LOGIN_TIMEOUT_MS / 1000}s for sign-in to complete`);
    try {
      await page.waitForURL(url => !isLoginUrl(url.toString()), { timeout: LOGIN_TIMEOUT_MS });
      log('sign-in complete');
    } catch {
      await context.close();
      throw new Error('Login timed out or browser was closed.');
    }
  }

  log('settling site page');
  await page.waitForLoadState('networkidle').catch(() => {});
  let spToken = capturedTokens.sp;
  if (!spToken) {
    log('no token captured on first pass — retrying networkidle navigation');
    try {
      await page.goto(siteUrl, { waitUntil: 'networkidle', timeout: 30000 });
      spToken = capturedTokens.sp;
    } catch {}
  }

  log('collecting tenant cookies');
  const allCookies = await context.cookies();
  const cookieStr = buildCookieString(allCookies, tenantHost);
  await context.close();

  if (!cookieStr) {
    throw new Error(`No cookies found for ${tenantUrl}. Run sp-api auth login --site ${site}`);
  }

  const authData = {
    SP_SITE: siteUrl,
    SP_COOKIES: cookieStr,
    ...(spToken && { SP_TOKEN: spToken }),
  };
  ensureDir(path.dirname(authFile));
  fs.writeFileSync(authFile, JSON.stringify(authData, null, 2) + '\n', { mode: 0o600 });
  setPrivateMode(authFile, 0o600);
  log(`wrote ${authFile} (token=${!!spToken})`);
  return { cookieStr, siteUrl, spToken };
}

function logout(options = {}) {
  const profileDir = options.profileDir || PROFILE_DIR;
  const authFile = options.authFile || AUTH_FILE;
  let cleared = false;
  if (fs.existsSync(profileDir)) {
    fs.rmSync(profileDir, { recursive: true, force: true });
    cleared = true;
  }
  if (fs.existsSync(authFile)) {
    fs.rmSync(authFile);
    cleared = true;
  }
  return { cleared, authFile, profileDir };
}

module.exports = {
  AUTH_FILE,
  PROFILE_DIR,
  DATA_DIR,
  authStatus,
  authenticate,
  buildCookieString,
  isLoginUrl,
  logout,
  parseSiteInput,
  readAuthFile,
};
