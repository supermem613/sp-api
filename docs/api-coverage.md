# API Coverage

`sp-api schema` is the source of truth for implemented commands. This document summarizes current and planned semantic coverage.

---

## Implemented Capability Groups

| Group | Verbs | Notes |
|-------|-------|-------|
| `auth` | `login`, `logout`, `status` | Playwright Edge persistent-context auth |
| `lists` | `list`, `get`, `fields`, `create`, `delete`, `add-field`, `items`, `add-item`, `update-item`, `delete-item` | List discovery, field schema, list-field creation with idempotent `--if-missing`, and core item CRUD |
| `files` | `list`, `get`, `folder`, `download`, `upload`, `delete`, `recycle`, `move`, `copy`, `create-folder`, `delete-folder`, `recycle-folder` | Folder/file operations for small text uploads, folder expansion, recycle, folder CRUD, and common file movement |
| `search` | `query` | Site-scoped SharePoint search with explicit query text and result shaping |
| `sites` | `get`, `discovery` | Site metadata and visible list/library discovery |
| `pages` | `list`, `get`, `checkout`, `save-fields`, `publish`, `discard-checkout` | Modern Site Pages read and lifecycle operations |
| `permissions` | `get` | Web-level role assignment inspection |
| `schema` | `show` | Machine-readable contract for all groups and verbs |
| `doctor` | `run` | Local health checks |

Inspect the live contract:

```bash
sp-api schema
sp-api schema lists add-item
sp-api schema files upload
sp-api schema pages publish
sp-api schema search query
```

## Planned Capability Groups

These are exposed as planned groups in `sp-api schema` but are not implemented as commands yet:

| Group | Planned verbs |
|-------|---------------|
| `search` | `people` |
| `pages` | `create`, `delete` |
| `users` | `list`, `get`, `me`, `ensure` |
| `permissions` | `grant`, `revoke`, `break-inheritance`, `reset-inheritance` |
| `sites` | `subsites`, `navigation` |

## Adding Coverage

Do not add raw HTTP examples as the public interface. To add coverage:

1. Add a semantic verb in `src/registry.js`.
2. Include params, endpoint metadata, examples, output docs, and auth requirements.
3. Add command/schema/help tests.
4. Update this coverage summary.

## Unsupported by Design

| Operation | Why | Alternative |
|-----------|-----|-------------|
| Raw REST passthrough | Would bypass the semantic command contract | Add a capability verb |
| Email sending | No SharePoint REST equivalent | Use Outlook tools |
| Teams messaging | No SharePoint REST equivalent | Use Teams tools |
| Enterprise RAG grounding | Requires separate orchestration | Use search/files plus local reasoning |
