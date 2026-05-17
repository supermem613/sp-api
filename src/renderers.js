'use strict';

const { capabilities, plannedCapabilities } = require('./registry');

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function emitSchema(targetCapability, targetVerb) {
  if (!targetCapability) {
    return {
      version: '0.1.0',
      capabilities: cloneJson(capabilities),
      plannedCapabilities: cloneJson(plannedCapabilities),
    };
  }
  const capability = capabilities[targetCapability];
  if (!capability) return null;
  if (!targetVerb) return cloneJson(capability);
  const verb = capability.verbs[targetVerb];
  if (!verb) return null;
  return cloneJson(verb);
}

function renderRootHelp() {
  const lines = [
    'sp-api - agentic SharePoint capability CLI',
    '',
    'Usage:',
    '  sp-api <capability> <verb> [options]',
    '  sp-api schema [capability] [verb]',
    '  sp-api doctor',
    '  sp-api update',
    '',
    'Capabilities:',
  ];
  for (const capability of Object.values(capabilities)) {
    lines.push(`  ${capability.id.padEnd(12)} ${capability.summary}`);
  }
  lines.push('');
  lines.push('Planned capability groups:');
  for (const [group, verbs] of Object.entries(plannedCapabilities)) {
    lines.push(`  ${group.padEnd(12)} ${verbs.join(', ')}`);
  }
  lines.push('');
  lines.push('Run "sp-api <capability> --help" or "sp-api schema <capability>" for generated details.');
  return lines.join('\n') + '\n';
}

function renderCapabilityHelp(capability) {
  const lines = [
    `sp-api ${capability.id} - ${capability.summary}`,
    '',
    capability.description || capability.summary,
    '',
    'Usage:',
    `  sp-api ${capability.id} <verb> [options]`,
    '',
    'Verbs:',
  ];
  for (const [verbName, verb] of Object.entries(capability.verbs)) {
    lines.push(`  ${verbName.padEnd(14)} ${verb.summary}`);
  }
  lines.push('');
  lines.push(`Run "sp-api ${capability.id} <verb> --help" for generated option details.`);
  return lines.join('\n') + '\n';
}

function renderVerbHelp(capability, verbName, verb) {
  const topLevelRun = verbName === 'run' && ['doctor', 'update'].includes(capability.id);
  const command = topLevelRun ? `sp-api ${capability.id}` : `sp-api ${capability.id} ${verbName}`;
  const lines = [
    `${command} - ${verb.summary}`,
    '',
    'Usage:',
    `  ${command}${verb.params.length ? ' [options]' : ''}`,
    '',
    `Auth: ${verb.auth}`,
    `Method: ${verb.method}`,
  ];
  if (verb.path) lines.push(`Endpoint: ${verb.path}`);
  if (verb.requiresDigest) lines.push('Requires digest: yes');
  if (verb.params.length) {
    lines.push('');
    lines.push('Options:');
    for (const param of verb.params) {
      const required = param.required ? 'required' : `optional${Object.hasOwn(param, 'default') ? `, default ${param.default}` : ''}`;
      lines.push(`  --${param.name.padEnd(14)} ${param.type.padEnd(8)} ${required}. ${param.doc}`);
    }
  }
  if (verb.examples.length) {
    lines.push('');
    lines.push('Examples:');
    for (const example of verb.examples) {
      lines.push(`  ${example}`);
    }
  }
  return lines.join('\n') + '\n';
}

function renderSkillRouter() {
  const commandRows = [];
  for (const capability of Object.values(capabilities)) {
    commandRows.push(`| \`${capability.id}\` | ${capability.summary} | \`sp-api ${capability.id} --help\` |`);
  }
  return `---
name: sp-api
description: "Use when you need to interact with SharePoint through the agentic sp-api CLI for lists, files, auth, schema inspection, and other SharePoint capabilities."
metadata:
  author: "Marcus Markiewicz"
  version: "1.0"
  license: "MIT"
  repo: "https://github.com/supermem613/sp-api"
---

# sp-api

This bundled skill is a thin router for the \`sp-api\` CLI. Use the CLI for SharePoint work. The CLI is agentic-only: stdout is JSON, progress and remediation go to stderr, and help/schema are generated from the same capability registry.

## Execution sequence

1. Run \`sp-api doctor\` if setup or auth is uncertain.
2. Use \`sp-api schema\` to inspect the full machine-readable contract, or \`sp-api schema <capability> <verb>\` for one command.
3. Run semantic commands such as \`sp-api lists items --title Tasks --top 25\` or \`sp-api files list --folder "/sites/team/Shared Documents"\`.
4. If a capability is not listed in \`schema\`, do not fall back to raw HTTP. Report the missing capability so a verb can be added.

## Capabilities

| Capability | Purpose | Details |
|------------|---------|---------|
${commandRows.join('\n')}

## References

Load these only when you need deeper SharePoint REST details behind a capability:

| File | Covers |
|------|--------|
| \`references/list-operations.md\` | List item CRUD, CAML queries, fields, views |
| \`references/file-operations.md\` | Upload, download, copy, move, versions, folders |
| \`references/search.md\` | SP Search and KQL syntax |
| \`references/site-discovery.md\` | Site properties, lists, fields, content types |
| \`references/page-operations.md\` | Site pages and news posts |
| \`references/user-permissions.md\` | Users, permissions, role assignments |
| \`references/advanced-operations.md\` | Recycle bin, navigation, features |
| \`references/api-patterns.md\` | OData, CAML, pagination, error handling |
`;
}

module.exports = {
  emitSchema,
  renderRootHelp,
  renderCapabilityHelp,
  renderVerbHelp,
  renderSkillRouter,
};
