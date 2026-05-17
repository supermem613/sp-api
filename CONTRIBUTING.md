# Contributing

Guide for engineers working on the `sp-api` CLI and its bundled skill router.

---

## Quick Start

```bash
git clone https://github.com/supermem613/sp-api
cd sp-api
npm install
npm run build
npm link
npm test
```

**Prerequisites:** Node.js 24+, Microsoft Edge for Playwright auth.

## Project Structure

```text
bin/
  sp-api.js                 # Package bin entrypoint
src/
  registry.js               # Single source of truth for capabilities, schema, help
  renderers.js              # Generated help, schema, and SKILL router renderers
  sp-api-core.js            # Agentic CLI dispatcher and JSON envelopes
  sharepoint-auth.js        # Playwright persistent-context auth
  sharepoint-fetch.js       # Shared fetch with retry and diagnostics
  sharepoint-rest.js        # SharePoint REST execution and request digest handling
.claude/skills/sp-api/
  SKILL.md                  # Lean agent router generated from registry
  references/               # Lazy-loaded REST background references
docs/
  AGENTIC_CONTRACT.md       # sp-api command and output contract
  architecture.md
  setup-guide.md
  api-coverage.md
  auth-deep-dive.md
evals/
  run-evals.md
tests/
  test-scripts.js           # Built-in auth, REST, and fetch module tests
  test-sp-api.js            # sp-api registry/schema/help/envelope tests
  test-integration.js       # Live SharePoint tests
```

## Architecture

This is a **CLI with a bundled skill**. The `sp-api` CLI is the product surface and owns all SharePoint behavior. The bundled `SKILL.md` is a thin router that tells agents to call semantic `sp-api` commands instead of composing raw HTTP.

The registry in `src/registry.js` is the source of truth for:

- command groups and verbs
- parameters and examples
- endpoint/method metadata
- `sp-api schema`
- generated `--help`
- generated `SKILL.md` router
- test expectations

Implementation logic belongs under `src/`. The skill directory should stay a lean router plus references.

## Development Workflow

### Running Tests

```bash
npm run build
npm test
npm run test:integration
```

`npm run build` validates that generated artifacts match the registry and that the `sp-api` bin is wired correctly. `npm run link:local` runs the build and then `npm link` for local CLI development.

Use `sp-api update` from linked or git-clone installs to self-update. It runs `git pull --ff-only`, skips install and build when already current, and otherwise runs `npm install --no-audit --no-fund` plus `npm run build`.

`npm test` must stay fast and offline. It validates the `sp-api` CLI contract and its built-in auth/REST modules.

### Linting the Skill

Run `lint-skill` after changing `SKILL.md` or reference files:

```bash
node C:\Users\marcusm\.copilot\skills\lint-skill\scripts\lint-skill.mjs --findings-only .claude\skills\sp-api
```

The linter should be clean before delivery. New errors or warnings must be fixed before shipping.

### Authenticating for Development

```bash
sp-api auth login --site contoso.sharepoint.com/sites/mysite
sp-api auth status
```

Use `--force` for interactive re-login and `sp-api auth logout` to clear the saved profile.

## Modifying Capabilities

Add or change commands in `src/registry.js` first. The registry change should drive schema, help, docs, and tests.

Rules:

1. **No raw passthrough.** Add semantic capability verbs instead of exposing arbitrary HTTP.
2. **Generated help.** Do not hand-write command help separate from the registry.
3. **Generated skill router.** `SKILL.md` must match `renderSkillRouter()`.
4. **Auth isolation.** Only `auth` may load Playwright. REST capability commands must not import Playwright directly or transitively.
5. **Agentic envelopes.** Non-help commands write one JSON object to stdout. Remediation goes to stderr.

## Modifying Reference Files

References provide SharePoint REST background for agents after they have selected a semantic `sp-api` capability. Prefer examples that start with `sp-api schema <capability> <verb>` or a semantic `sp-api` command.

When adding a new operation:

1. Add a semantic verb to `src/registry.js`.
2. Add or update tests in `tests/test-sp-api.js`.
3. Update `docs/api-coverage.md`.
4. Update references only for background details the schema cannot express.

## Adding Evals

Evals should verify that agents choose `sp-api`, pass arguments that match `sp-api schema`, and interpret the JSON envelope.

All live test data must be prefixed with `SP_API_EVAL_`.

## Code Style

- `'use strict'` at the top of CommonJS scripts
- Shebang line on executable bins
- JSON stdout for non-help commands
- stderr for progress and remediation
- Comments explain why, not mechanics
