#!/usr/bin/env node
'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { existsSync, readFileSync, mkdtempSync, rmSync, writeFileSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join } = require('node:path');
const { pathToFileURL } = require('node:url');
const { authenticate, authStatus, buildCookieString, logout, parseSiteInput, readAuthFile } = require('../src/sharepoint-auth');
const { executeSharePointRequest, loadAuth, parseResponseBody } = require('../src/sharepoint-rest');
const { extractCode, formatError, spFetch } = require('../src/sharepoint-fetch');

describe('sp-api source ownership', () => {
  it('keeps implementation logic in src instead of the skill directory', () => {
    assert.ok(existsSync(join(__dirname, '..', 'src', 'sharepoint-auth.js')));
    assert.ok(existsSync(join(__dirname, '..', 'src', 'sharepoint-rest.js')));
    assert.ok(existsSync(join(__dirname, '..', 'src', 'sharepoint-fetch.js')));
    assert.strictEqual(existsSync(join(__dirname, '..', '.claude', 'skills', 'sp-api', 'scripts')), false);
  });

  it('keeps doctor output focused on sp-api modules', () => {
    const core = readFileSync(join(__dirname, '..', 'src', 'sp-api-core.js'), 'utf8');
    assert.doesNotMatch(core, /Legacy/);
  });
});

describe('SharePoint auth module', () => {
  it('parses SharePoint site input', () => {
    assert.deepStrictEqual(parseSiteInput('https://contoso.sharepoint.com/sites/team/'), {
      tenantHost: 'contoso.sharepoint.com',
      sitePath: '/sites/team',
    });
  });

  it('builds a cookie string from tenant cookies', () => {
    const cookies = [
      { name: 'FedAuth', value: 'a', domain: '.contoso.sharepoint.com' },
      { name: 'rtFa', value: 'b', domain: '.sharepoint.com' },
      { name: 'Other', value: 'c', domain: '.example.com' },
    ];
    assert.strictEqual(buildCookieString(cookies, 'contoso.sharepoint.com'), 'FedAuth=a; rtFa=b');
  });

  it('reports auth status from the auth file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'sp-api-auth-'));
    try {
      const authFile = join(dir, 'auth.json');
      writeFileSync(authFile, JSON.stringify({ SP_SITE: 'https://contoso.sharepoint.com/sites/team', SP_COOKIES: 'FedAuth=a' }));
      assert.deepStrictEqual(authStatus(authFile), {
        authFile,
        exists: true,
        site: 'https://contoso.sharepoint.com/sites/team',
        hasCookies: true,
        hasToken: false,
      });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('logs out by clearing auth state paths', () => {
    const dir = mkdtempSync(join(tmpdir(), 'sp-api-logout-'));
    try {
      const profileDir = join(dir, 'profile');
      const authFile = join(dir, 'auth.json');
      writeFileSync(authFile, '{}');
      writeFileSync(join(dir, 'profile-marker'), '');
      require('node:fs').mkdirSync(profileDir);
      const result = logout({ profileDir, authFile });
      assert.strictEqual(result.cleared, true);
      assert.strictEqual(existsSync(authFile), false);
      assert.strictEqual(existsSync(profileDir), false);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('authenticates with injectable Playwright and writes auth.json', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'sp-api-playwright-'));
    try {
      const authFile = join(dir, 'auth.json');
      const page = {
        on: () => {},
        goto: async () => {},
        url: () => 'https://contoso.sharepoint.com/sites/team',
        waitForLoadState: async () => {},
      };
      const context = {
        pages: () => [page],
        cookies: async () => [{ name: 'FedAuth', value: 'a', domain: '.contoso.sharepoint.com' }],
        close: async () => {},
      };
      const playwright = {
        chromium: {
          launchPersistentContext: async () => context,
        },
      };
      const result = await authenticate('contoso.sharepoint.com/sites/team', { playwright, authFile, profileDir: join(dir, 'profile') });
      assert.strictEqual(result.siteUrl, 'https://contoso.sharepoint.com/sites/team');
      assert.strictEqual(readAuthFile(authFile).SP_COOKIES, 'FedAuth=a');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('SharePoint REST module', () => {
  it('loads auth from auth.json and rejects missing credentials', () => {
    const dir = mkdtempSync(join(tmpdir(), 'sp-api-rest-'));
    try {
      const authFile = join(dir, 'auth.json');
      assert.throws(() => loadAuth(authFile), /Run sp-api auth login/);
      writeFileSync(authFile, JSON.stringify({ SP_SITE: 'https://contoso.sharepoint.com/sites/team' }));
      assert.throws(() => loadAuth(authFile), /credentials/);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('parses JSON, text, and empty response bodies', () => {
    assert.deepStrictEqual(parseResponseBody('{"value":1}'), { value: 1 });
    assert.strictEqual(parseResponseBody('plain'), 'plain');
    assert.strictEqual(parseResponseBody(''), null);
  });

  it('executes GET requests through the built-in REST client', async () => {
    const calls = [];
    const result = await executeSharePointRequest(
      { method: 'GET' },
      '_api/web',
      '',
      {
        auth: { SP_SITE: 'https://contoso.sharepoint.com/sites/team', SP_COOKIES: 'FedAuth=a' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          return { ok: true, status: 200, text: async () => '{"Title":"Team"}' };
        },
      },
    );
    assert.strictEqual(result.data.Title, 'Team');
    assert.strictEqual(calls[0].url, 'https://contoso.sharepoint.com/sites/team/_api/web');
    assert.strictEqual(calls[0].options.method, 'GET');
  });

  it('fetches a digest and uses method override for mutations', async () => {
    const calls = [];
    const result = await executeSharePointRequest(
      { method: 'PATCH' },
      '_api/web/lists',
      '{"Title":"New"}',
      {
        auth: { SP_SITE: 'https://contoso.sharepoint.com/sites/team', SP_TOKEN: 'token' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          if (url.endsWith('/_api/contextinfo')) {
            return { ok: true, status: 200, text: async () => '{"FormDigestValue":"digest"}' };
          }
          return { ok: true, status: 204, text: async () => '' };
        },
      },
    );
    assert.strictEqual(result.data, null);
    assert.strictEqual(calls[1].options.method, 'POST');
    assert.strictEqual(calls[1].options.headers['X-HTTP-Method'], 'PATCH');
    assert.strictEqual(calls[1].options.headers['If-Match'], '*');
  });

  it('honors per-verb REST accept and content type overrides', async () => {
    const calls = [];
    await executeSharePointRequest(
      { method: 'POST', accept: 'application/json;odata=verbose', contentType: 'application/json;odata=nometadata' },
      '_api/web/example',
      '{"Title":"Example"}',
      {
        auth: { SP_SITE: 'https://contoso.sharepoint.com/sites/team', SP_TOKEN: 'token' },
        fetch: async (url, options) => {
          calls.push({ url, options });
          if (url.endsWith('/_api/contextinfo')) {
            return { ok: true, status: 200, text: async () => '{"FormDigestValue":"digest"}' };
          }
          return { ok: true, status: 200, text: async () => '{"ok":true}' };
        },
      },
    );
    assert.strictEqual(calls[1].options.headers.Accept, 'application/json;odata=verbose');
    assert.strictEqual(calls[1].options.headers['Content-Type'], 'application/json;odata=nometadata');
  });
});

describe('SharePoint fetch module', () => {
  it('walks nested error causes', () => {
    assert.strictEqual(extractCode({ cause: { code: 'ETIMEDOUT' } }), 'ETIMEDOUT');
  });

  it('formats actionable network errors', () => {
    assert.match(formatError({ message: 'failed', code: 'ENOTFOUND' }), /Hint: DNS lookup failed/);
  });

  it('retries retryable fetch failures without shelling out', async () => {
    let attempts = 0;
    const response = await spFetch(pathToFileURL(__filename).toString(), {}, {
      fetch: async () => {
        attempts++;
        if (attempts === 1) {
          const err = new Error('timeout');
          err.code = 'ETIMEDOUT';
          throw err;
        }
        return { ok: true };
      },
    });
    assert.strictEqual(response.ok, true);
    assert.strictEqual(attempts, 2);
  });
});
