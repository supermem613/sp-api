---
name: sp-api
description: "Use when you need to interact with SharePoint through the agentic sp-api CLI for lists, files, auth, schema inspection, and other SharePoint capabilities."
metadata:
  author: "Marcus Markiewicz"
  version: "1.0"
  license: "MIT"
  repo: "https://github.com/supermem613/sp-api"
---

# sp-api

This bundled skill is a thin router for the `sp-api` CLI. Use the CLI for SharePoint work. The CLI is agentic-only: stdout is JSON, progress and remediation go to stderr, and help/schema are generated from the same capability registry.

## Execution sequence

1. Run `sp-api doctor` if setup or auth is uncertain.
2. Use `sp-api schema` to inspect the full machine-readable contract, or `sp-api schema <capability> <verb>` for one command.
3. Run semantic commands such as `sp-api lists items --title Tasks --top 25` or `sp-api files list --folder "/sites/team/Shared Documents"`.
4. If a capability is not listed in `schema`, do not fall back to raw HTTP. Report the missing capability so a verb can be added.

## Capabilities

| Capability | Purpose | Details |
|------------|---------|---------|
| `auth` | Manage SharePoint browser-session authentication | `sp-api auth --help` |
| `lists` | Work with SharePoint lists and list items | `sp-api lists --help` |
| `files` | Work with SharePoint files and folders | `sp-api files --help` |
| `search` | Search SharePoint content | `sp-api search --help` |
| `sites` | Discover SharePoint site metadata | `sp-api sites --help` |
| `pages` | Work with SharePoint site pages | `sp-api pages --help` |
| `permissions` | Inspect SharePoint permissions | `sp-api permissions --help` |
| `schema` | Inspect the generated SharePoint capability schema | `sp-api schema --help` |
| `doctor` | Run local health checks for the agentic CLI | `sp-api doctor --help` |
| `update` | Self-update this sp-api checkout | `sp-api update --help` |

## References

Load these only when you need deeper SharePoint REST details behind a capability:

| File | Covers |
|------|--------|
| `references/list-operations.md` | List item CRUD, CAML queries, fields, views |
| `references/file-operations.md` | Upload, download, copy, move, versions, folders |
| `references/search.md` | SP Search and KQL syntax |
| `references/site-discovery.md` | Site properties, lists, fields, content types |
| `references/page-operations.md` | Site pages and news posts |
| `references/user-permissions.md` | Users, permissions, role assignments |
| `references/advanced-operations.md` | Recycle bin, navigation, features |
| `references/api-patterns.md` | OData, CAML, pagination, error handling |
