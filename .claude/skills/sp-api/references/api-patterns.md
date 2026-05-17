# sp-api Patterns

`sp-api` hides SharePoint REST mechanics behind semantic capability commands. Agents should inspect schema and call capabilities rather than constructing raw endpoints.

## Schema-first usage

```bash
sp-api schema
sp-api schema lists items
sp-api schema files upload
```

The schema returns:

- capability and verb ids
- required and optional parameters
- underlying REST method and endpoint metadata
- auth and digest requirements
- examples
- JSON envelope shape

## Output envelope

Non-help commands return one JSON object on stdout:

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

Failures still use JSON stdout:

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

Progress and remediation go to stderr.

## Auth and digest

Use `sp-api auth login --site <site>` for browser-session auth. Mutating commands carry digest requirements in schema and handle digest through internal helpers.

## Missing coverage

There is no raw HTTP passthrough. If a SharePoint operation is missing, add a semantic capability verb to `src/registry.js`, add tests, and update coverage docs.

## OData and CAML

OData and CAML remain SharePoint implementation details. Use them only when adding or extending a semantic verb. Agents should not emit ad hoc raw REST calls as the public workflow.
