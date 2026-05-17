# Advanced SharePoint Operations

Advanced operations are not implemented as `sp-api` commands yet. Inspect the planned surface with:

```bash
sp-api schema
```

Candidate future capability groups or verbs include:

| Area | Candidate capability |
|------|----------------------|
| Recycle bin | `sites recycle-bin`, `sites restore-recycle-bin` |
| Navigation | `sites navigation` |
| Site features | `sites features` |
| Webhooks | `lists subscriptions` |
| Retention labels | `lists retention-label` |
| Approvals | `lists approvals` |

Add advanced coverage only as semantic verbs in `src/registry.js` with tests and schema. Do not document raw REST commands as the agent-facing path.
