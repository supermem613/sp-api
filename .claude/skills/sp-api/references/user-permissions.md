# SharePoint Users and Permissions

`permissions get` is implemented for web-level role assignment inspection. `users` remains planned. Inspect the current contract with:

```bash
sp-api schema permissions
sp-api permissions get
```

Planned `users` verbs:

| Verb | Purpose |
|------|---------|
| `me` | Get current user |
| `list` | List site users |
| `get` | Get one user |
| `ensure` | Ensure/resolve a user |

Implemented `permissions` verbs:

| Verb | Purpose |
|------|---------|
| `get` | Read web role assignments expanded with members and role bindings |

Planned mutation verbs:

| Verb | Purpose |
|------|---------|
| `grant` | Grant a role |
| `revoke` | Remove a role |
| `break-inheritance` | Break permission inheritance |
| `reset-inheritance` | Restore inherited permissions |

Prompt pattern:

> Show who has access to this site and what roles they have.

Use:

```bash
sp-api permissions get
```

Permission changes are remote mutations. When implemented, commands must use explicit params, JSON envelopes, and clear failure messages. Do not use raw REST as a fallback.
