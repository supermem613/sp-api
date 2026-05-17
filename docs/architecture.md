## Architecture & Design Decisions

### What This Is

This is a SharePoint Online CLI with a bundled skill router. The `sp-api` CLI is the product surface: it maps capability verbs to SharePoint REST implementation details and returns stable JSON envelopes. Agents may load the bundled `SKILL.md`, but that skill only routes them to the CLI.

### Why a CLI with a Bundled Skill Instead of an MCP Server

- **Agentic command surface** — agents call `sp-api lists items`, `sp-api files upload`, and `sp-api schema`, not raw HTTP helpers.
- **No long-running server** — commands are short-lived and work in normal shells.
- **Cross-platform** — Node.js and Playwright work on Windows, macOS, and Linux.
- **Generated contract** — schema, help, and the skill router come from one registry.
- **Auth isolation** — Playwright is limited to `sp-api auth`; REST commands stay on the lightweight hot path.

### Command Architecture

```text
Agent loads SKILL.md
    |
    v
sp-api <capability> <verb>
    |
    +-- src/registry.js        capability specs, params, examples, endpoints
    +-- src/renderers.js       generated help, schema, SKILL router
    +-- src/sp-api-core.js     dispatcher and JSON envelopes
    +-- src/sharepoint-auth.js Playwright auth and auth state
    +-- src/sharepoint-rest.js SharePoint REST execution
    +-- src/sharepoint-fetch.js retry and diagnostics
    |
    v
SharePoint Online
```

The CLI owns the implementation. The skill directory contains only the router and lazy-loaded references.

### Auth Architecture

```text
sp-api auth login
    |
    v
src/sharepoint-auth.js
    |
    v
Playwright persistent Edge context
    |
    +-- profile: ~/.sp-api/browser-profile/
    +-- auth:    ~/.sp-api/auth.json
```

Only the auth path loads Playwright. Tests assert that REST capability files do not import Playwright.

### Bundled Skill Architecture

```text
Agent loads SKILL.md
    |
    +-- command model
    +-- schema/help routing
    +-- reference file index
    |
    v
Agent calls sp-api schema <capability> <verb>
    |
    v
Agent calls semantic command
```

Reference files provide background REST details only after a capability has been selected. They should not route agents around `sp-api`.

### Design Decisions Log

| Decision | Chosen | Why |
|----------|--------|-----|
| Product form | CLI with bundled skill | Tested CLI owns behavior. Skill is a thin router |
| Primary surface | `sp-api` semantic commands | Capability-oriented and agentic |
| Source of truth | `src/registry.js` | Schema/help/SKILL/tests cannot drift |
| Auth method | Playwright persistent context | No app registration, browser-equivalent access |
| Auth boundary | `sp-api auth` only | Prevents Playwright from entering REST hot path |
| Raw HTTP fallback | Not supported | Missing coverage should become a semantic verb |
| Output | JSON envelope | Stable machine-readable agent contract |
| Self-update | `sp-api update` | Git-clone installs can pull, install, and rebuild in one command |

### What This CLI Cannot Do

| Capability | Why | Workaround |
|-----------|-----|-----------|
| Raw arbitrary HTTP passthrough | Deliberately excluded to keep the CLI semantic | Add a capability verb |
| Email sending | No SharePoint REST equivalent | Use Outlook or other mail tools |
| Teams messaging | No SharePoint REST equivalent | Use Teams tools |
| Sharing links | No current semantic command | Add a capability when supported |
| Enterprise-wide M365 search | Current scope is SharePoint-site-oriented | Use M365 or SharePoint admin tools |
| Server-side code execution | SharePoint sandboxing | Run code locally |
