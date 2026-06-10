# sp-api Agentic Contract

`sp-api` is the stable SharePoint command surface. The repo includes a bundled skill router, but agents should treat the CLI as the product API and call semantic SharePoint capability commands instead of composing raw HTTP verbs.

## Source of truth

The capability registry in `src/registry.js` defines every command, option, endpoint, help string, schema entry, example, and output contract. Help and schema must be generated from that registry. Do not hand-write separate command help.

## Output contract

Non-help commands write one JSON object to stdout:

```json
{
  "ok": true,
  "command": "lists.items",
  "data": {},
  "error": null,
  "meta": {
    "schemaVersion": "0.1.0"
  }
}
```

Failures keep stdout machine-readable:

```json
{
  "ok": false,
  "command": "lists.items",
  "data": null,
  "error": {
    "code": "VALIDATION_FAILED",
    "message": "Missing required option --title"
  },
  "meta": {
    "schemaVersion": "0.1.0"
  }
}
```

Progress, browser-login instructions, and remediation details go to stderr. Help text is the only human-oriented stdout output.

## Command model

Commands are SharePoint capability groups with bounded verbs:

```text
sp-api auth login|logout|status
sp-api lists list|get|fields|create|delete|add-field|items|add-item|update-item|delete-item
sp-api files list|get|download|upload|create-folder|delete-folder|recycle-folder|delete|move|copy|folder|recycle
sp-api search query
sp-api sites get|discovery
sp-api pages list|get|checkout|save-fields|publish|discard-checkout
sp-api permissions get
sp-api schema [capability] [verb]
sp-api doctor
sp-api update
```

There is no raw HTTP passthrough. If a SharePoint action is missing, add a semantic verb to the registry with tests.

## Auth isolation

Only `sp-api auth` may load Playwright. REST capability commands must stay on the built-in REST client path and must not import Playwright directly or transitively.

## Self-update

`sp-api update` is for git-clone installs. It runs `git pull --ff-only`, skips install and build when already current, and runs `npm install --no-audit --no-fund` plus `npm run build` when changes arrive. It still returns the standard JSON envelope on stdout.

## Mutation safety

Mutating commands must make required inputs explicit and return structured failures. Future remote mutations should add preview/apply semantics where that materially reduces risk.
