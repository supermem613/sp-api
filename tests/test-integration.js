#!/usr/bin/env node
'use strict';

const { describe, it, before } = require('node:test');
const assert = require('node:assert');
const { execFileSync } = require('node:child_process');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { AUTH_FILE } = require('../src/sharepoint-auth');

const repoRoot = join(__dirname, '..');
const cliPath = join(repoRoot, 'bin', 'sp-api.js');
const siteArg = process.argv.find(a => a.includes('sharepoint') && a.includes('.com') && !a.startsWith('-'));

function runSpApi(args, options = {}) {
  const stdout = execFileSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: 'utf8',
    timeout: options.timeout || 120000,
    stdio: ['pipe', 'pipe', options.inheritStderr ? 'inherit' : 'pipe'],
  });
  return JSON.parse(stdout);
}

function cachedSite() {
  try {
    return JSON.parse(readFileSync(AUTH_FILE, 'utf8')).SP_SITE || '';
  } catch {
    return '';
  }
}

const targetSite = siteArg || cachedSite();
if (!targetSite) {
  console.log('Skipping integration tests because no SharePoint site is available.');
  console.log('Pass one with: npm run test:integration -- contoso.sharepoint.com/sites/mysite');
  process.exit(1);
}

describe('sp-api live SharePoint integration', () => {
  before(() => {
    if (siteArg) {
      runSpApi(['auth', 'login', '--site', siteArg], { timeout: 300000, inheritStderr: true });
    }
  });

  it('reports auth status through sp-api', () => {
    const status = runSpApi(['auth', 'status']);
    assert.strictEqual(status.ok, true);
    assert.strictEqual(status.data.exists, true);
    assert.ok(status.data.site);
  });

  it('runs doctor without helper script checks', () => {
    const doctor = runSpApi(['doctor']);
    assert.strictEqual(doctor.ok, true);
    assert.ok(doctor.data.checks.some(check => check.name === 'auth-file'));
    assert.ok(!doctor.data.checks.some(check => check.name.includes('.js')));
  });

  it('can list SharePoint lists through the semantic command surface', () => {
    const lists = runSpApi(['lists', 'list']);
    assert.strictEqual(lists.ok, true);
    assert.ok(lists.data);
  });
});
