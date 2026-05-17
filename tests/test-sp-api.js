#!/usr/bin/env node
'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { spawnSync } = require('node:child_process');
const { existsSync, readFileSync } = require('node:fs');
const { join } = require('node:path');
const { capabilities } = require('../src/registry');
const {
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  renderSkillRouter,
} = require('../src/renderers');
const { gitPullMadeNoChanges, selfUpdate } = require('../src/sp-api-core');

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
    assert.strictEqual(pkg.version, '1.1.0');
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
