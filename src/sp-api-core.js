'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { capabilities } = require('./registry');
const { emitSchema, renderRootHelp, renderCapabilityHelp, renderVerbHelp } = require('./renderers');
const { AUTH_FILE, authenticate, authStatus, logout } = require('./sharepoint-auth');
const { executeSharePointRequest } = require('./sharepoint-rest');
const { buildCreateFieldXmlBody } = require('./list-fields');

const repoRoot = path.join(__dirname, '..');

const BODY_BUILDERS = {
  'create-field-xml': buildCreateFieldXmlBody,
};

const PRE_CHECKS = {
  'field-exists': fieldExistsPreCheck,
};

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
  if (value === true && param.type !== 'boolean') {
    throw new Error(`--${param.name} requires a value`);
  }
  if (param.type === 'file-text') {
    return fs.readFileSync(String(value), 'utf8');
  }
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
    if (value !== undefined && param.enum && !param.enum.includes(value)) {
      throw new Error(`--${param.name} must be one of: ${param.enum.join(', ')}`);
    }
    if (value !== undefined) values[param.mapsTo || param.name] = value;
  }
  for (const group of spec.requiresOneOf || []) {
    if (!group.some(name => flags[name] !== undefined)) {
      throw new Error(`Missing one of required options: ${group.map(name => `--${name}`).join(' or ')}`);
    }
  }
  for (const rule of spec.requiresWhen || []) {
    if (!ruleTriggerMatches(rule.when, flags)) continue;
    for (const requiredName of rule.requires) {
      if (flags[requiredName] === undefined) {
        const trigger = rule.when.value === undefined
          ? `--${rule.when.param} is set`
          : `--${rule.when.param} is ${rule.when.value}`;
        throw new Error(`Missing required option --${requiredName} when ${trigger}`);
      }
    }
  }
  return values;
}

function ruleTriggerMatches(when, flags) {
  const raw = flags[when.param];
  if (when.value === undefined) return raw !== undefined;
  if (raw === when.value) return true;
  return String(raw) === String(when.value);
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
  if (spec.bodyBuilder) {
    const builder = BODY_BUILDERS[spec.bodyBuilder.kind];
    if (!builder) throw new Error(`Unknown bodyBuilder kind: ${spec.bodyBuilder.kind}`);
    return builder(values);
  }
  if (Object.hasOwn(values, 'body')) return JSON.stringify(values.body);
  if (spec.bodyParam) return String(values[spec.bodyParam] || '');
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
  if (spec.preCheck) {
    const preChecker = PRE_CHECKS[spec.preCheck.kind];
    if (preChecker) {
      const skip = await preChecker(spec.preCheck, values);
      if (skip) return skip;
    }
  }
  const { endpoint, body } = buildSharePointRequest(spec, values);
  try {
    return await executeSharePointRequest(spec, endpoint, body);
  } catch (err) {
    if (spec.notFoundOkParam && values[spec.notFoundOkParam] === true && /HTTP 404\b/.test(err.message)) {
      return {
        data: { missing: true },
        endpoint,
        method: spec.method || 'GET',
      };
    }
    throw err;
  }
}

async function fieldExistsPreCheck(preCheck, values, deps = {}) {
  if (!values[preCheck.whenParam]) return null;
  const listTitle = values[preCheck.listParam];
  const internalName = values[preCheck.nameParam];
  const endpoint =
    `_api/web/lists/getbytitle('${encodeSharePointValue(listTitle)}')` +
    `/fields/getbyinternalnameortitle('${encodeSharePointValue(internalName)}')` +
    `?$select=InternalName,Title,TypeAsString`;
  try {
    const existing = await executeSharePointRequest({ method: 'GET' }, endpoint, '', deps);
    return {
      data: {
        skipped: true,
        reason: 'already-exists',
        name: internalName,
        field: existing.data,
      },
      endpoint,
      method: 'GET',
    };
  } catch (err) {
    if (isFieldMissingError(err)) return null;
    throw err;
  }
}

function isFieldMissingError(err) {
  const message = err?.message || '';
  if (/HTTP 404\b/.test(message)) return true;
  if (/HTTP 400\b/.test(message) && /Column\b[^\n]*\bdoes not exist/i.test(message)) return true;
  return false;
}

function buildSharePointRequest(spec, values) {
  return {
    endpoint: addQuery(replacePlaceholders(spec.path, values), spec.query, values),
    body: buildBody(spec, values),
  };
}

function doctor() {
  const nodeMajor = Number.parseInt(process.versions.node.split('.')[0], 10);
  const raw = [
    { name: 'node', ok: nodeMajor >= 24, detail: process.version, hint: 'Install Node.js 24 or later' },
    { name: 'auth-file', ok: fs.existsSync(AUTH_FILE), detail: AUTH_FILE, hint: 'Run "sp-api auth login --site <site>" to create it' },
  ];
  const checks = raw.map(c => (c.ok ? { name: c.name, ok: true, detail: c.detail } : c));
  return { checks };
}

function gitPullMadeNoChanges(output) {
  return /already up[- ]to[- ]date\.?/i.test(output);
}

function runCommand(command, args, cwd, deps = {}) {
  const options = {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  };
  const nodeExecutable = deps.nodeExecutable || process.execPath;
  const resolveNpm = deps.resolveNpmCliPath || (() => resolveNpmCliPath(nodeExecutable));

  if (process.platform === 'win32' && /^npm(?:\.cmd)?$/i.test(command)) {
    const npmCli = resolveNpm();
    if (npmCli) return spawnSync(nodeExecutable, [npmCli, ...args], options);
  }
  return spawnSync(command, args, options);
}

function resolveNpmCliPath(nodeExecutable = process.execPath) {
  const nodeDir = path.dirname(nodeExecutable);
  const candidates = [
    path.join(nodeDir, 'node_modules', 'npm', 'bin', 'npm-cli.js'),
    path.join(nodeDir, '..', 'node_modules', 'npm', 'bin', 'npm-cli.js')
  ];
  for (const candidate of candidates) {
    try {
      if (fs.statSync(candidate).isFile()) return candidate;
    } catch {}
  }
  return null;
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
    const data = await authenticate(flags.site, { forceLogin: !!flags.force, verbose: !!flags.verbose });
    return { ok: true, data: { site: data.siteUrl, authenticated: true, hasToken: !!data.spToken } };
  }
  const data = logout();
  return { ok: true, data: { loggedOut: data.cleared, authFile: data.authFile, profileDir: data.profileDir } };
}

function fail(stdout, code, command, message, details) {
  writeJson(stdout, envelope(false, command, null, { code, message, details }));
}

function writeDownloadToOut(values, execution) {
  if (values.out === true) {
    throw new Error('Missing value for --out');
  }
  const outValue = String(values.out);
  if (!outValue.trim()) {
    throw new Error('--out must be a non-empty path');
  }
  const outPath = path.resolve(outValue);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  const bodyBuffer = execution.bodyBuffer || Buffer.from(typeof execution.data === 'string' ? execution.data : '', 'utf8');
  fs.writeFileSync(outPath, bodyBuffer);
  return {
    path: outPath,
    bytes: bodyBuffer.length,
  };
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
    const failed = data.checks.filter(check => !check.ok).map(check => check.name);
    const ok = failed.length === 0;
    const result = {
      ok,
      command: 'doctor',
      data,
      error: ok
        ? null
        : {
            code: 'DOCTOR_FAILED',
            message: `Doctor check failed: ${failed.join(', ')}. Run "sp-api auth login --site <site>" if auth-file is the only failure.`,
            failed,
          },
      meta: { schemaVersion: '0.1.0' },
    };
    writeJson(stdout, result);
    exit(ok ? 0 : 1);
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
  let data = execution.data;
  if (spec.id === 'files.download' && values.out) {
    try {
      data = writeDownloadToOut(values, execution);
    } catch (err) {
      fail(stdout, 'WRITE_FAILED', spec.id, `Could not write --out file: ${err.message}`);
      exit(1);
      return;
    }
  }
  writeJson(stdout, envelope(true, spec.id, data, null, { endpoint: execution.endpoint, method: execution.method }));
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
  buildSharePointRequest,
  runSharePoint,
  selfUpdate,
  runCommand,
  writeDownloadToOut,
  fieldExistsPreCheck,
};
