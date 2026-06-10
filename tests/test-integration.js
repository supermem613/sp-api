#!/usr/bin/env node
'use strict';

const { describe, it, before } = require('node:test');
const assert = require('node:assert');
const { execFileSync } = require('node:child_process');
const { existsSync, mkdtempSync, readFileSync, rmSync } = require('node:fs');
const { tmpdir } = require('node:os');
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

  it('downloads a file to --out and preserves local bytes', t => {
    const discovery = runSpApi(['sites', 'discovery']);
    assert.strictEqual(discovery.ok, true);
    const libraries = (discovery.data?.value || []).filter(item => item?.BaseType === 1 && item?.RootFolder?.ServerRelativeUrl);
    if (!libraries.length) {
      t.skip('No visible document library found for files integration test');
      return;
    }

    const folder = libraries[0].RootFolder.ServerRelativeUrl;
    const suffix = Date.now();
    const name = `SP_API_INTEGRATION_DOWNLOAD_${suffix}.txt`;
    const serverPath = `${folder}/${name}`;
    const content = `integration-download-${suffix}`;
    const tmp = mkdtempSync(join(tmpdir(), 'sp-api-download-'));
    const outPath = join(tmp, 'downloaded.txt');

    try {
      const upload = runSpApi(['files', 'upload', '--folder', folder, '--name', name, '--content', content]);
      assert.strictEqual(upload.ok, true);

      const download = runSpApi(['files', 'download', '--path', serverPath, '--out', outPath]);
      assert.strictEqual(download.ok, true);
      assert.strictEqual(download.data.path, outPath);
      assert.strictEqual(download.data.bytes, Buffer.byteLength(content, 'utf8'));
      assert.strictEqual(existsSync(outPath), true);
      assert.strictEqual(readFileSync(outPath, 'utf8'), content);
    } finally {
      rmSync(tmp, { recursive: true, force: true });
      try {
        runSpApi(['files', 'delete', '--path', serverPath]);
      } catch {
        // Best-effort cleanup; a failed delete should not hide download assertion results.
      }
    }
  });

  it('creates text, datetime, and choice list fields and is idempotent under --if-missing', () => {
    const suffix = Date.now();
    const listTitle = `SP_API_INTEGRATION_FIELDS_${suffix}`;

    const create = runSpApi(['lists', 'create', '--title', listTitle]);
    assert.strictEqual(create.ok, true, `lists create failed: ${JSON.stringify(create.error)}`);

    try {
      const owner = runSpApi(['lists', 'add-field', '--title', listTitle, '--name', 'Owner', '--type', 'text', '--if-missing', 'true']);
      assert.strictEqual(owner.ok, true, `add-field Owner failed: ${JSON.stringify(owner.error)}`);
      assert.notStrictEqual(owner.data?.skipped, true, 'first Owner creation should not be skipped');

      const reviewDate = runSpApi([
        'lists', 'add-field',
        '--title', listTitle,
        '--name', 'ReviewDate',
        '--display-name', 'Review Date',
        '--type', 'datetime',
        '--format', 'date-only',
        '--if-missing', 'true',
      ]);
      assert.strictEqual(reviewDate.ok, true, `add-field ReviewDate failed: ${JSON.stringify(reviewDate.error)}`);

      const status = runSpApi([
        'lists', 'add-field',
        '--title', listTitle,
        '--name', 'Status',
        '--type', 'choice',
        '--choices', 'Not Started,In Progress,Completed',
        '--if-missing', 'true',
      ]);
      assert.strictEqual(status.ok, true, `add-field Status failed: ${JSON.stringify(status.error)}`);

      const priority = runSpApi([
        'lists', 'add-field',
        '--title', listTitle,
        '--name', 'Priority',
        '--type', 'choice',
        '--choices', 'High,Normal,Low',
        '--required', 'true',
        '--if-missing', 'true',
      ]);
      assert.strictEqual(priority.ok, true, `add-field Priority failed: ${JSON.stringify(priority.error)}`);

      const fields = runSpApi(['lists', 'fields', '--title', listTitle]);
      assert.strictEqual(fields.ok, true);
      const internalNames = new Set((fields.data?.value || []).map(field => field.InternalName));
      for (const expected of ['Owner', 'ReviewDate', 'Status', 'Priority']) {
        assert.ok(internalNames.has(expected), `field ${expected} should exist on the list`);
      }

      const replay = runSpApi([
        'lists', 'add-field',
        '--title', listTitle,
        '--name', 'Status',
        '--type', 'choice',
        '--choices', 'Not Started,In Progress,Completed',
        '--if-missing', 'true',
      ]);
      assert.strictEqual(replay.ok, true);
      assert.strictEqual(replay.data?.skipped, true, '--if-missing true should short-circuit when the field exists');
      assert.strictEqual(replay.data?.reason, 'already-exists');
      assert.strictEqual(replay.data?.name, 'Status');
    } finally {
      try {
        runSpApi(['lists', 'delete', '--title', listTitle]);
      } catch {
        // Best-effort cleanup; a failed delete should not hide field assertion results.
      }
    }
  });
});
