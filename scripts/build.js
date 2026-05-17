#!/usr/bin/env node
'use strict';

const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { renderSkillRouter } = require('../src/renderers');
const { capabilities } = require('../src/registry');

const repoRoot = path.join(__dirname, '..');
const packageJsonPath = path.join(repoRoot, 'package.json');
const skillPath = path.join(repoRoot, '.claude', 'skills', 'sp-api', 'SKILL.md');
const binPath = path.join(repoRoot, 'bin', 'sp-api.js');

function normalizeNewlines(text) {
  return text.replace(/\r\n/g, '\n');
}

function main() {
  const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  assert.strictEqual(pkg.version, '1.2.0', 'package version should be 1.2.0');
  assert.strictEqual(pkg.bin?.['sp-api'], 'bin/sp-api.js', 'package bin must expose sp-api');

  const bin = fs.readFileSync(binPath, 'utf8');
  assert.match(bin.split(/\r?\n/)[0], /node/, 'sp-api bin must have a node shebang');

  const skill = normalizeNewlines(fs.readFileSync(skillPath, 'utf8'));
  assert.strictEqual(skill, renderSkillRouter(), 'SKILL.md must match the registry-rendered router');

  for (const capability of ['auth', 'lists', 'files', 'search', 'sites', 'pages', 'permissions', 'schema', 'doctor', 'update']) {
    assert.ok(capabilities[capability], `missing capability: ${capability}`);
  }

  process.stdout.write('Build validation passed.\n');
}

main();
