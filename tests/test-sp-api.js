#!/usr/bin/env node
'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { spawnSync } = require('node:child_process');
const { existsSync, readFileSync, writeFileSync, mkdtempSync, rmSync } = require('node:fs');
const { tmpdir } = require('node:os');
const { join } = require('node:path');
const { capabilities } = require('../src/registry');
const {
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  renderSkillRouter,
} = require('../src/renderers');
const { buildSharePointRequest, collectParams, gitPullMadeNoChanges, runCommand, selfUpdate, writeDownloadToOut } = require('../src/sp-api-core');

const repoRoot = join(__dirname, '..');
const cliPath = join(repoRoot, 'bin', 'sp-api.js');
const skillPath = join(repoRoot, '.claude', 'skills', 'sp-api', 'SKILL.md');

function runCli(args) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function parseJson(stdout) {
  return JSON.parse(stdout);
}

describe('sp-api package wiring', () => {
  it('exposes a bin script with a shebang', () => {
    const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8'));
    assert.match(pkg.version, /^\d+\.\d+\.\d+(?:-[\w.]+)?$/);
    assert.strictEqual(pkg.bin['sp-api'], 'bin/sp-api.js');
    assert.strictEqual(pkg.scripts.build, 'node scripts/build.js');
    assert.strictEqual(pkg.scripts.prepare, 'npm run build');
    assert.strictEqual(pkg.scripts['link:local'], 'npm run build && npm link');
    assert.strictEqual(pkg.scripts.auth, 'sp-api auth status');
    assert.strictEqual(pkg.scripts.login, 'sp-api auth login');
    assert.strictEqual(pkg.scripts.logout, 'sp-api auth logout');
    assert.match(readFileSync(cliPath, 'utf8').split(/\r?\n/)[0], /node/);
  });
});

describe('registry invariants', () => {
  it('has unique capability and verb ids with examples and output contracts', () => {
    const ids = new Set();
    for (const [capabilityName, capability] of Object.entries(capabilities)) {
      assert.strictEqual(capabilityName, capability.id);
      assert.ok(capability.summary);
      for (const [verbName, verb] of Object.entries(capability.verbs)) {
        assert.strictEqual(verb.id, `${capabilityName}.${verbName}`);
        assert.ok(!ids.has(verb.id), `Duplicate id ${verb.id}`);
        ids.add(verb.id);
        assert.ok(verb.summary, `${verb.id} missing summary`);
        assert.ok(verb.output?.envelope, `${verb.id} missing envelope`);
        assert.ok(verb.examples?.length, `${verb.id} missing examples`);
        const params = new Set();
        for (const param of verb.params) {
          assert.match(param.name, /^[a-z][a-z0-9-]*$/);
          assert.ok(!params.has(param.name), `${verb.id} duplicate param ${param.name}`);
          params.add(param.name);
          assert.ok(param.doc, `${verb.id}.${param.name} missing docs`);
        }
      }
    }
  });

  it('does not expose raw HTTP verbs as top-level capabilities', () => {
    assert.ok(!capabilities.get);
    assert.ok(!capabilities.post);
    assert.ok(!capabilities.request);
  });

  it('exposes compatible semantic capabilities brought over from spfs', () => {
    assert.ok(capabilities.search.verbs.query);
    assert.ok(capabilities.sites.verbs.get);
    assert.ok(capabilities.sites.verbs.discovery);
    assert.ok(capabilities.pages.verbs.list);
    assert.ok(capabilities.pages.verbs.get);
    assert.ok(capabilities.pages.verbs.checkout);
    assert.ok(capabilities.pages.verbs['save-fields']);
    assert.ok(capabilities.pages.verbs.publish);
    assert.ok(capabilities.pages.verbs['discard-checkout']);
    assert.ok(capabilities.permissions.verbs.get);
    assert.ok(capabilities.lists.verbs.fields);
    assert.ok(capabilities.files.verbs.folder);
    assert.ok(capabilities.files.verbs.recycle);
    assert.ok(capabilities.files.verbs['create-folder']);
    assert.ok(capabilities.files.verbs['delete-folder']);
    assert.ok(capabilities.files.verbs['recycle-folder']);
  });

  it('keeps shipped verbs out of the planned capability lists', () => {
    const schema = emitSchema();
    assert.deepStrictEqual(schema.plannedCapabilities.search, ['people']);
    assert.deepStrictEqual(schema.plannedCapabilities.pages, ['create', 'delete']);
    assert.deepStrictEqual(schema.plannedCapabilities.permissions, ['grant', 'revoke', 'break-inheritance', 'reset-inheritance']);
    assert.deepStrictEqual(schema.plannedCapabilities.sites, ['subsites', 'navigation']);
  });

  it('keeps Playwright isolated to the auth module', () => {
    const nonAuthFiles = [
      join(repoRoot, 'bin', 'sp-api.js'),
      join(repoRoot, 'src', 'registry.js'),
      join(repoRoot, 'src', 'renderers.js'),
      join(repoRoot, 'src', 'sp-api-core.js'),
      join(repoRoot, 'src', 'sharepoint-rest.js'),
      join(repoRoot, 'src', 'sharepoint-fetch.js'),
    ];
    for (const file of nonAuthFiles) {
      assert.doesNotMatch(readFileSync(file, 'utf8'), /require\(['"]playwright['"]\)/, file);
    }
    assert.match(readFileSync(join(repoRoot, 'src', 'sharepoint-auth.js'), 'utf8'), /require\(['"]playwright['"]\)/);
  });
});

describe('schema output', () => {
  it('emits a full schema envelope', () => {
    const r = runCli(['schema']);
    assert.strictEqual(r.status, 0);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, true);
    assert.strictEqual(json.command, 'schema');
    assert.deepStrictEqual(json.data.capabilities.lists.verbs.items, capabilities.lists.verbs.items);
  });

  it('emits a focused capability and verb schema', () => {
    const cap = parseJson(runCli(['schema', 'lists']).stdout);
    assert.strictEqual(cap.data.id, 'lists');
    const verb = parseJson(runCli(['schema', 'lists', 'add-item']).stdout);
    assert.strictEqual(verb.data.id, 'lists.add-item');
    assert.strictEqual(verb.data.method, 'POST');
  });

  it('emits focused schemas for newly implemented capability groups', () => {
    const search = parseJson(runCli(['schema', 'search', 'query']).stdout);
    assert.strictEqual(search.data.id, 'search.query');
    const page = parseJson(runCli(['schema', 'pages', 'save-fields']).stdout);
    assert.strictEqual(page.data.params.find(param => param.name === 'body').type, 'json');
    const permissions = parseJson(runCli(['schema', 'permissions', 'get']).stdout);
    assert.strictEqual(permissions.data.path, '_api/web/roleassignments');
  });

  it('fails unknown schema targets with a JSON envelope', () => {
    const r = runCli(['schema', 'missing']);
    assert.strictEqual(r.status, 2);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, false);
    assert.strictEqual(json.error.code, 'UNKNOWN_SCHEMA');
  });
});

describe('generated help', () => {
  it('generates root help from the registry', () => {
    const r = runCli(['--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderRootHelp());
    assert.match(r.stdout, /lists\s+Work with SharePoint lists/);
  });

  it('generates capability help from the registry', () => {
    const r = runCli(['lists', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderCapabilityHelp(capabilities.lists));
    assert.match(r.stdout, /add-item/);
  });

  it('generates help for new capability groups from the registry', () => {
    const pages = runCli(['pages', '--help']);
    assert.strictEqual(pages.status, 0);
    assert.strictEqual(pages.stdout, renderCapabilityHelp(capabilities.pages));
    assert.match(pages.stdout, /save-fields/);

    const search = runCli(['search', 'query', '--help']);
    assert.strictEqual(search.status, 0);
    assert.strictEqual(search.stdout, renderVerbHelp(capabilities.search, 'query', capabilities.search.verbs.query));
    assert.match(search.stdout, /--row-limit/);
  });

  it('generates verb help from the registry', () => {
    const r = runCli(['lists', 'items', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderVerbHelp(capabilities.lists, 'items', capabilities.lists.verbs.items));
    assert.match(r.stdout, /--title/);
    assert.match(r.stdout, /Endpoint:/);
  });
});

describe('JSON envelope behavior', () => {
  it('returns validation failures as JSON on stdout', () => {
    const r = runCli(['lists', 'items']);
    assert.strictEqual(r.status, 2);
    assert.strictEqual(r.stderr, '');
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, false);
    assert.strictEqual(json.command, 'lists.items');
    assert.strictEqual(json.error.code, 'VALIDATION_FAILED');
    assert.match(json.error.message, /--title/);
  });

  it('rejects valueless string options before running SharePoint requests', () => {
    const r = runCli(['files', 'download', '--path', '/sites/team/Shared Documents/file.txt', '--out']);
    assert.strictEqual(r.status, 2);
    assert.strictEqual(r.stderr, '');
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, false);
    assert.strictEqual(json.command, 'files.download');
    assert.strictEqual(json.error.code, 'VALIDATION_FAILED');
    assert.match(json.error.message, /--out requires a value/);
  });

  it('rejects unknown capabilities and verbs without raw fallback', () => {
    const cap = parseJson(runCli(['request', '--method', 'GET']).stdout);
    assert.strictEqual(cap.ok, false);
    assert.strictEqual(cap.error.code, 'UNKNOWN_CAPABILITY');
    const verb = parseJson(runCli(['lists', 'request']).stdout);
    assert.strictEqual(verb.ok, false);
    assert.strictEqual(verb.error.code, 'UNKNOWN_VERB');
  });

  it('reports auth status without loading Playwright', () => {
    const r = runCli(['auth', 'status']);
    assert.strictEqual(r.status, 0);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.ok, true);
    assert.ok(Object.hasOwn(json.data, 'exists'));
    assert.ok(json.data.authFile.endsWith(join('.sp-api', 'auth.json')));
  });

  it('keeps the doctor checks payload on both success and failure', () => {
    const r = runCli(['doctor']);
    const json = parseJson(r.stdout);
    assert.strictEqual(json.command, 'doctor');
    assert.ok(Array.isArray(json.data.checks), 'doctor.data.checks must always be present');
    assert.ok(json.data.checks.length >= 2, 'doctor should report at least node + auth-file checks');
    const knownNames = new Set(['node', 'auth-file']);
    for (const check of json.data.checks) {
      assert.ok(knownNames.has(check.name), `unexpected check name: ${check.name}`);
      assert.strictEqual(typeof check.ok, 'boolean');
      if (check.ok) {
        assert.strictEqual(check.hint, undefined, `passing checks must not carry a hint (${check.name})`);
      } else {
        assert.strictEqual(typeof check.hint, 'string');
      }
    }
    if (json.ok) {
      assert.strictEqual(r.status, 0);
      assert.strictEqual(json.error, null);
    } else {
      assert.strictEqual(r.status, 1);
      assert.strictEqual(json.error.code, 'DOCTOR_FAILED');
      assert.ok(Array.isArray(json.error.failed), 'error.failed must list the failed check names');
      assert.ok(json.error.failed.length >= 1);
      for (const name of json.error.failed) {
        const check = json.data.checks.find(c => c.name === name);
        assert.ok(check, `failed check ${name} must appear in data.checks`);
        assert.strictEqual(check.ok, false);
      }
      assert.match(json.error.message, /Doctor check failed/);
    }
  });
});

describe('SharePoint request construction', () => {
  it('builds encoded search query URLs without exposing raw HTTP args', () => {
    const request = buildSharePointRequest(capabilities.search.verbs.query, {
      query: "Title:'Roadmap'",
      'row-limit': 10,
      'select-properties': 'Title,Path',
    });
    assert.strictEqual(
      request.endpoint,
      "_api/search/query?querytext='Title%3A''Roadmap'''&rowlimit=10&selectproperties=Title%2CPath",
    );
    assert.strictEqual(request.body, '');
  });

  it('builds page field update bodies from explicit JSON without OData string escaping', () => {
    const request = buildSharePointRequest(capabilities.pages.verbs['save-fields'], {
      'item-id': 42,
      body: { Title: "Team's page" },
    });
    assert.strictEqual(request.endpoint, "_api/web/lists/getbytitle('Site Pages')/items(42)");
    assert.strictEqual(request.body, '{"Title":"Team\'s page"}');
  });

  it('builds folder, recycle, permissions, and site discovery endpoints', () => {
    assert.strictEqual(
      buildSharePointRequest(capabilities.files.verbs.folder, { folder: '/sites/team/Shared Documents' }).endpoint,
      "_api/web/getfolderbyserverrelativeurl('/sites/team/Shared Documents')?$expand=Folders%2CFiles",
    );
    assert.strictEqual(
      buildSharePointRequest(capabilities.files.verbs.recycle, { path: '/sites/team/Shared Documents/old.txt' }).endpoint,
      "_api/web/getfilebyserverrelativeurl('/sites/team/Shared Documents/old.txt')/recycle",
    );
    assert.deepStrictEqual(
      buildSharePointRequest(capabilities.files.verbs['create-folder'], { path: "/sites/team/Shared Documents/Agent's Skills" }),
      {
        endpoint: '_api/web/folders',
        body: '{"__metadata":{"type":"SP.Folder"},"ServerRelativeUrl":"/sites/team/Shared Documents/Agent\'s Skills"}',
      },
    );
    assert.strictEqual(
      buildSharePointRequest(capabilities.files.verbs['delete-folder'], { path: "/sites/team/Shared Documents/Agent's Skills", 'missing-ok': false }).endpoint,
      "_api/web/getfolderbyserverrelativeurl('/sites/team/Shared Documents/Agent''s Skills')",
    );
    assert.strictEqual(
      buildSharePointRequest(capabilities.files.verbs['recycle-folder'], { path: '/sites/team/Shared Documents/Old Folder' }).endpoint,
      "_api/web/getfolderbyserverrelativeurl('/sites/team/Shared Documents/Old Folder')/recycle",
    );
    assert.strictEqual(
      buildSharePointRequest(capabilities.permissions.verbs.get, {}).endpoint,
      '_api/web/roleassignments?$expand=Member%2CRoleDefinitionBindings',
    );
    assert.strictEqual(
      buildSharePointRequest(capabilities.sites.verbs.discovery, {}).endpoint,
      '_api/web/lists?$filter=Hidden eq false&$select=Id%2CTitle%2CBaseTemplate%2CBaseType%2CItemCount%2CRootFolder%2FServerRelativeUrl&$expand=RootFolder',
    );
  });

  it('builds upload bodies from inline content or content files', () => {
    const tmp = mkdtempSync(join(tmpdir(), 'sp-api-upload-'));
    try {
      const bodyPath = join(tmp, 'SKILL.md');
      writeFileSync(bodyPath, 'hello from file', 'utf8');
      assert.strictEqual(
        buildSharePointRequest(capabilities.files.verbs.upload, {
          folder: '/sites/team/Shared Documents',
          name: 'notes.txt',
          content: 'hello inline',
          overwrite: true,
        }).body,
        'hello inline',
      );
      const values = collectParams(capabilities.files.verbs.upload, {
        folder: '/sites/team/Shared Documents',
        name: 'SKILL.md',
        'content-file': bodyPath,
      });
      assert.strictEqual(values.content, 'hello from file');
      assert.strictEqual(buildSharePointRequest(capabilities.files.verbs.upload, values).body, 'hello from file');
    } finally {
      rmSync(tmp, { recursive: true, force: true });
    }
  });

  it('requires upload content or content-file', () => {
    assert.throws(
      () => collectParams(capabilities.files.verbs.upload, {
        folder: '/sites/team/Shared Documents',
        name: 'SKILL.md',
      }),
      /--content or --content-file/,
    );
  });

  it('supports files.download --out and writes binary-safe bytes', () => {
    const values = collectParams(capabilities.files.verbs.download, {
      path: '/sites/team/Shared Documents/image.png',
      out: 'downloads/image.png',
    });
    assert.strictEqual(values.path, '/sites/team/Shared Documents/image.png');
    assert.strictEqual(values.out, 'downloads/image.png');

    const tmp = mkdtempSync(join(tmpdir(), 'sp-api-download-'));
    const originalCwd = process.cwd();
    try {
      process.chdir(tmp);
      const metadata = writeDownloadToOut(values, { bodyBuffer: Buffer.from([0, 255, 127, 65]) });
      assert.ok(existsSync(metadata.path));
      assert.strictEqual(metadata.bytes, 4);
      assert.deepStrictEqual([...readFileSync(metadata.path)], [0, 255, 127, 65]);
    } finally {
      process.chdir(originalCwd);
      rmSync(tmp, { recursive: true, force: true });
    }
  });
});

describe('SKILL.md router generation', () => {
  it('matches the registry-rendered router', () => {
    assert.strictEqual(readFileSync(skillPath, 'utf8').replace(/\r\n/g, '\n'), renderSkillRouter());
  });

  it('routes every implemented capability', () => {
    const content = readFileSync(skillPath, 'utf8');
    for (const capability of Object.keys(capabilities)) {
      assert.match(content, new RegExp(`sp-api ${capability}`));
    }
  });
});

describe('update command', () => {
  it('appears in root help and schema', () => {
    const help = runCli(['--help']).stdout;
    assert.match(help, /update\s+Self-update this sp-api checkout/);
    const schema = parseJson(runCli(['schema', 'update', 'run']).stdout);
    assert.strictEqual(schema.data.id, 'update.run');
  });

  it('generates update help from the registry', () => {
    const r = runCli(['update', '--help']);
    assert.strictEqual(r.status, 0);
    assert.strictEqual(r.stdout, renderVerbHelp(capabilities.update, 'run', capabilities.update.verbs.run));
  });

  it('recognizes current and older no-change git pull output', () => {
    assert.strictEqual(gitPullMadeNoChanges('Already up to date.'), true);
    assert.strictEqual(gitPullMadeNoChanges('Already up-to-date.'), true);
    assert.strictEqual(gitPullMadeNoChanges('Fast-forward\n package.json | 2 +-'), false);
  });

  it('skips install and build when git pull made no changes', () => {
    const commands = [];
    const result = selfUpdate({
      repoRoot,
      isGitRepo: () => true,
      runCommand: (command, args) => {
        commands.push(`${command} ${args.join(' ')}`);
        return { status: 0, stdout: 'Already up to date.\n', stderr: '' };
      },
    });
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.data.updated, false);
    assert.deepStrictEqual(commands, ['git pull --ff-only']);
  });

  it('runs install and build when git pull returns changes', () => {
    const commands = [];
    const result = selfUpdate({
      repoRoot,
      isGitRepo: () => true,
      runCommand: (command, args) => {
        commands.push(`${command} ${args.join(' ')}`);
        return { status: 0, stdout: command === 'git' ? 'Fast-forward\n' : 'ok\n', stderr: '' };
      },
    });
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.data.updated, true);
    assert.deepStrictEqual(commands, [
      'git pull --ff-only',
      'npm install --no-audit --no-fund',
      'npm run build',
    ]);
  });

  it('executes npm via node + npm-cli.js on Windows', () => {
    const tmp = mkdtempSync(join(tmpdir(), 'sp-api-npm-cli-'));
    const fakeNpmCli = join(tmp, 'npm-cli.js');
    writeFileSync(
      fakeNpmCli,
      [
        "'use strict';",
        'process.stdout.write(JSON.stringify({',
        '  argv0: process.argv[0],',
        '  args: process.argv.slice(2),',
        '}));',
      ].join('\n'),
      'utf8',
    );

    const platformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');

    try {
      Object.defineProperty(process, 'platform', { value: 'win32' });
      const result = runCommand('npm', ['run', 'build'], repoRoot, {
        resolveNpmCliPath: () => fakeNpmCli,
      });
      assert.strictEqual(result.status, 0);
      const payload = JSON.parse(result.stdout);
      assert.strictEqual(payload.argv0, process.execPath);
      assert.deepStrictEqual(payload.args, ['run', 'build']);
    } finally {
      Object.defineProperty(process, 'platform', platformDescriptor);
      rmSync(tmp, { recursive: true, force: true });
    }
  });
});

describe('agentic contract documentation', () => {
  it('documents schema-generated help and no raw passthrough', () => {
    const doc = readFileSync(join(repoRoot, 'docs', 'AGENTIC_CONTRACT.md'), 'utf8');
    assert.match(doc, /generated from that registry/);
    assert.match(doc, /There is no raw HTTP passthrough/);
  });

  it('keeps required repo files present', () => {
    assert.ok(existsSync(join(repoRoot, 'src', 'registry.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'renderers.js')));
    assert.ok(existsSync(join(repoRoot, 'src', 'sp-api-core.js')));
  });
});
