'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { capabilities } = require('./registry');
const { emitSchema, renderRootHelp, renderCapabilityHelp, renderVerbHelp } = require('./renderers');
const { AUTH_FILE, authenticate, authStatus, logout } = require('./sharepoint-auth');
const { executeSharePointRequest } = require('./sharepoint-rest');

const repoRoot = path.join(__dirname, '..');

function envelope(ok, command, data, error, meta = {}) {
  return {
    ok,
    command,
    data: ok ? data : null,
    error: ok ? null : error,
    meta: { ...meta, schemaVersion: '0.1.0' },
  };
}

function writeJson(stdout, value) {
  stdout.write(JSON.stringify(value, null, 2) + '\n');
}

function parseArgs(args) {
  const positional = [];
  const flags = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }
    const name = arg.slice(2);
    const next = args[i + 1];
    if (!next || next.startsWith('--')) {
      flags[name] = true;
      continue;
    }
    flags[name] = next;
    i++;
  }
  return { positional, flags };
}

function coerceValue(param, value) {
  if (value === undefined && Object.hasOwn(param, 'default')) return param.default;
  if (value === undefined) return undefined;
  if (param.type === 'number') {
    const numberValue = Number(value);
    if (!Number.isFinite(numberValue)) throw new Error(`--${param.name} must be a number`);
    return numberValue;
  }
  if (param.type === 'boolean') {
    if (value === true) return true;
    if (value === 'true') return true;
    if (value === 'false') return false;
    throw new Error(`--${param.name} must be true or false`);
  }
  if (param.type === 'json') {
    try {
      return JSON.parse(value);
    } catch {
      throw new Error(`--${param.name} must be valid JSON`);
    }
  }
  return String(value);
}

function collectParams(spec, flags) {
  const values = {};
  for (const param of spec.params) {
    const value = coerceValue(param, flags[param.name]);
    if (value === undefined && param.required) {
      throw new Error(`Missing required option --${param.name}`);
    }
    if (value !== undefined) values[param.name] = value;
  }
  return values;
}

function replacePlaceholders(template, values) {
  return template.replace(/\{([^}]+)\}/g, (_, name) => {
    if (!Object.hasOwn(values, name)) return '';
    return encodeSharePointValue(values[name]);
  });
}

function encodeSharePointValue(value) {
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value).replace(/'/g, "''");
}

function addQuery(endpoint, query, values) {
  if (!query) return endpoint;
  const parts = [];
  for (const [name, template] of Object.entries(query)) {
    const raw = replacePlaceholders(template, values);
    if (raw === '') continue;
    parts.push(`${name}=${encodeURIComponent(raw)}`);
  }
  if (!parts.length) return endpoint;
  return `${endpoint}${endpoint.includes('?') ? '&' : '?'}${parts.join('&')}`;
}

function buildBody(spec, values) {
  if (Object.hasOwn(values, 'body')) return JSON.stringify(values.body);
  if (!spec.bodyTemplate) return '';
  function visit(value) {
    if (typeof value === 'string') {
      const match = value.match(/^\{([^}]+)\}$/);
      if (match) return values[match[1]];
      return replacePlaceholders(value, values);
    }
    if (Array.isArray(value)) return value.map(visit);
    if (value && typeof value === 'object') {
      const result = {};
      for (const [key, child] of Object.entries(value)) result[key] = visit(child);
      return result;
    }
    return value;
  }
  return JSON.stringify(visit(spec.bodyTemplate));
}

async function runSharePoint(spec, values) {
  const endpoint = addQuery(replacePlaceholders(spec.path, values), spec.query, values);
  const body = buildBody(spec, values);
  return executeSharePointRequest(spec, endpoint, body);
}

function doctor() {
  const checks = [
    { name: 'node', ok: true, detail: process.version, hint: 'Install Node.js 24 or later' },
    { name: 'cli-entrypoint', ok: fs.existsSync(path.join(repoRoot, 'bin', 'sp-api.js')), detail: 'sp-api executable', hint: 'Restore bin/sp-api.js' },
    { name: 'capability-registry', ok: fs.existsSync(path.join(repoRoot, 'src', 'registry.js')), detail: 'Generated schema and help source', hint: 'Restore src/registry.js' },
    { name: 'rest-client', ok: fs.existsSync(path.join(repoRoot, 'src', 'sharepoint-rest.js')), detail: 'SharePoint REST client', hint: 'Restore src/sharepoint-rest.js' },
    { name: 'auth-module', ok: fs.existsSync(path.join(repoRoot, 'src', 'sharepoint-auth.js')), detail: 'SharePoint auth module', hint: 'Restore src/sharepoint-auth.js' },
    { name: 'auth-file', ok: fs.existsSync(AUTH_FILE), detail: AUTH_FILE, hint: 'Run sp-api auth login --site <site>' },
  ];
  return { ok: checks.every(check => check.ok), checks };
}

function gitPullMadeNoChanges(output) {
  return /already up[- ]to[- ]date\.?/i.test(output);
}

function runCommand(command, args, cwd) {
  return spawnSync(command, args, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    shell: process.platform === 'win32',
  });
}

function isGitRepo(cwd) {
  const result = runCommand('git', ['rev-parse', '--is-inside-work-tree'], cwd);
  return result.status === 0 && result.stdout.trim() === 'true';
}

function selfUpdate(deps = {}) {
  const root = deps.repoRoot || repoRoot;
  const checkGitRepo = deps.isGitRepo || isGitRepo;
  const run = deps.runCommand || runCommand;
  const steps = [];

  if (!checkGitRepo(root)) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'NOT_GIT_REPO', message: 'sp-api update requires a git clone install' },
    };
  }

  const pull = run('git', ['pull', '--ff-only'], root);
  const pullOutput = `${pull.stdout || ''}${pull.stderr || ''}`.trim();
  steps.push({ name: 'git pull --ff-only', ok: pull.status === 0, output: pullOutput });
  if (pull.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'GIT_PULL_FAILED', message: pullOutput || 'git pull --ff-only failed' },
    };
  }
  if (gitPullMadeNoChanges(pullOutput)) {
    return { ok: true, data: { repoRoot: root, updated: false, steps } };
  }

  const install = run('npm', ['install', '--no-audit', '--no-fund'], root);
  const installOutput = `${install.stdout || ''}${install.stderr || ''}`.trim();
  steps.push({ name: 'npm install --no-audit --no-fund', ok: install.status === 0, output: installOutput });
  if (install.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'NPM_INSTALL_FAILED', message: installOutput || 'npm install failed' },
    };
  }

  const build = run('npm', ['run', 'build'], root);
  const buildOutput = `${build.stdout || ''}${build.stderr || ''}`.trim();
  steps.push({ name: 'npm run build', ok: build.status === 0, output: buildOutput });
  if (build.status !== 0) {
    return {
      ok: false,
      data: { repoRoot: root, steps },
      error: { code: 'BUILD_FAILED', message: buildOutput || 'npm run build failed' },
    };
  }

  return { ok: true, data: { repoRoot: root, updated: true, steps } };
}

async function runAuth(verbName, flags) {
  if (verbName === 'status') return { ok: true, data: authStatus() };
  if (verbName === 'login') {
    const data = await authenticate(flags.site, { forceLogin: !!flags.force });
    return { ok: true, data: { site: data.siteUrl, authenticated: true, hasToken: !!data.spToken } };
  }
  const data = logout();
  return { ok: true, data: { loggedOut: data.cleared, authFile: data.authFile, profileDir: data.profileDir } };
}

function fail(stdout, code, command, message, details) {
  writeJson(stdout, envelope(false, command, null, { code, message, details }));
}

async function main(args, io) {
  const { stdout, stderr, exit } = io;
  const parsed = parseArgs(args);
  const [capabilityName, verbName] = parsed.positional;
  if (!capabilityName || capabilityName === 'help') {
    stdout.write(renderRootHelp());
    exit(0);
    return;
  }
  if (capabilityName === 'schema') {
    const schema = emitSchema(verbName, parsed.positional[2]);
    if (!schema) {
      fail(stdout, 'UNKNOWN_SCHEMA', 'schema', `Unknown schema target: ${parsed.positional.slice(1).join(' ')}`);
      exit(2);
      return;
    }
    writeJson(stdout, envelope(true, 'schema', schema, null));
    exit(0);
    return;
  }
  if (capabilityName === 'doctor') {
    if (verbName === '--help' || parsed.flags.help) {
      stdout.write(renderVerbHelp(capabilities.doctor, 'run', capabilities.doctor.verbs.run));
      exit(0);
      return;
    }
    const data = doctor();
    writeJson(stdout, envelope(data.ok, 'doctor', data, data.ok ? null : { code: 'DOCTOR_FAILED', message: 'One or more checks failed' }));
    exit(data.ok ? 0 : 1);
    return;
  }
  if (capabilityName === 'update') {
    if (verbName === '--help' || parsed.flags.help) {
      stdout.write(renderVerbHelp(capabilities.update, 'run', capabilities.update.verbs.run));
      exit(0);
      return;
    }
    const result = selfUpdate();
    writeJson(stdout, envelope(result.ok, 'update.run', result.data, result.error));
    exit(result.ok ? 0 : 1);
    return;
  }
  const capability = capabilities[capabilityName];
  if (!capability) {
    fail(stdout, 'UNKNOWN_CAPABILITY', capabilityName, `Unknown capability: ${capabilityName}`);
    exit(2);
    return;
  }
  if (!verbName) {
    stdout.write(renderCapabilityHelp(capability));
    exit(0);
    return;
  }
  const spec = capability.verbs[verbName];
  if (!spec) {
    fail(stdout, 'UNKNOWN_VERB', `${capabilityName}.${verbName}`, `Unknown verb: ${capabilityName} ${verbName}`);
    exit(2);
    return;
  }
  if (parsed.flags.help) {
    stdout.write(renderVerbHelp(capability, verbName, spec));
    exit(0);
    return;
  }
  let values;
  try {
    values = collectParams(spec, parsed.flags);
  } catch (err) {
    fail(stdout, 'VALIDATION_FAILED', spec.id, err.message);
    exit(2);
    return;
  }
  if (capabilityName === 'auth') {
    try {
      const result = await runAuth(verbName, values);
      writeJson(stdout, envelope(result.ok, spec.id, result.data, null));
      exit(0);
    } catch (err) {
      writeJson(stdout, envelope(false, spec.id, null, { code: 'AUTH_FAILED', message: err.message }));
      exit(1);
    }
    return;
  }
  let execution;
  try {
    execution = await runSharePoint(spec, values);
  } catch (err) {
    fail(stdout, 'SHAREPOINT_REQUEST_FAILED', spec.id, err.message);
    exit(1);
    return;
  }
  writeJson(stdout, envelope(true, spec.id, execution.data, null, { endpoint: execution.endpoint, method: execution.method }));
  exit(0);
}

module.exports = {
  main,
  parseArgs,
  collectParams,
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  envelope,
  gitPullMadeNoChanges,
  runSharePoint,
  selfUpdate,
};
