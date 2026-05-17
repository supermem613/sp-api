# SharePoint Users and Permissions

`users` and `permissions` are planned `sp-api` capability groups. Inspect planned capabilities with:

```bash
sp-api schema
```

Planned `users` verbs:

| Verb | Purpose |
|------|---------|
| `me` | Get current user |
| `list` | List site users |
| `get` | Get one user |
| `ensure` | Ensure/resolve a user |

Planned `permissions` verbs:

| Verb | Purpose |
|------|---------|
| `get` | Read role assignments or permissions |
| `grant` | Grant a role |
| `revoke` | Remove a role |
| `break-inheritance` | Break permission inheritance |
| `reset-inheritance` | Restore inherited permissions |

Permission changes are remote mutations. When implemented, commands must use explicit params, JSON envelopes, and clear failure messages. Do not use raw REST as a fallback.
