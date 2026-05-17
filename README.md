# sp-api

**An agentic SharePoint CLI with a bundled thin skill for Claude Code, GitHub Copilot CLI, and other coding agents.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## What It Does

This repo is first and foremost the `sp-api` CLI. It includes a bundled skill only as a thin router so agents know when and how to call the CLI. Agents use semantic SharePoint capability commands instead of raw HTTP verbs. The CLI owns auth, command routing, generated help, generated schema, JSON envelopes, and the current lists/files implementation.

`sp-api` is agentic by default:

- JSON stdout for non-help commands
- Progress and remediation on stderr
- `sp-api schema` as the machine-readable source of truth
- Help generated from the same capability registry as schema
- No raw HTTP passthrough

## Current Command Surface

```text
sp-api auth   login | logout | status
sp-api lists  list | get | create | delete | items | add-item | update-item | delete-item
sp-api files  list | get | download | upload | delete | move | copy
sp-api schema [capability] [verb]
sp-api doctor
sp-api update
```

Planned capability groups are exposed in `sp-api schema` so agents can see what is not implemented yet without falling back to raw HTTP.

## Quick Start

```bash
git clone https://github.com/supermem613/sp-api
cd sp-api
npm install
npm run build
npm link
sp-api doctor
```

Authenticate once:

```bash
sp-api auth login --site contoso.sharepoint.com/sites/mysite
```

Then use semantic commands:

```bash
sp-api lists list
sp-api lists items --title Tasks --select Title,Id,Status --top 25
sp-api files list --folder "/sites/mysite/Shared Documents"
sp-api schema lists add-item
```

## Bundled Skill

The skill under `.claude/skills/sp-api` is not the product surface. It is a generated router plus lazy-loaded references. Install the CLI first, then install or copy the skill so agents route SharePoint tasks to `sp-api`.

### Claude Code

```claude
/install supermem613/sp-api
```

### Copilot CLI / Other Agents

Copy `.claude/skills/sp-api` into the agent's skill directory and install the package dependencies from this repo. The skill routes agents to `sp-api`.

## Auth

`sp-api auth login` uses Playwright with Microsoft Edge persistent context. First run may open Edge for interactive login. Subsequent runs use the saved browser profile headlessly.

- Browser profile: `~/.sp-api/browser-profile/`
- Auth file: `~/.sp-api/auth.json`
- Force re-login: `sp-api auth login --site <site> --force`
- Clear auth: `sp-api auth logout`

No app registration, client ID, tenant config, or secret is required.

## Tests

```bash
npm run build
npm test
npm run test:integration
```

`npm run build` validates generated artifacts and the `sp-api` bin before local linking or publishing. `npm link` and `npm run link:local` are supported for local development. For linked or git-clone installs, `sp-api update` pulls with `git pull --ff-only`, skips install/build when already current, and otherwise runs `npm install --no-audit --no-fund` plus `npm run build`.

`npm test` covers the `sp-api` registry, schema generation, help generation, JSON envelopes, SharePoint auth/REST internals, no raw fallback, auth isolation, package bin wiring, and SKILL router sync.

`npm run test:integration` is the live SharePoint test suite and requires an authenticated site.

## Docs

- [`docs/AGENTIC_CONTRACT.md`](docs/AGENTIC_CONTRACT.md) — `sp-api` stdout/stderr, schema, help, and command contract
- [`docs/setup-guide.md`](docs/setup-guide.md) — install and auth
- [`docs/architecture.md`](docs/architecture.md) — registry, CLI, and auth architecture
- [`docs/api-coverage.md`](docs/api-coverage.md) — current and planned capability coverage
- [`docs/auth-deep-dive.md`](docs/auth-deep-dive.md) — Playwright persistent-context auth details
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — development workflow

## License

[MIT](LICENSE)
