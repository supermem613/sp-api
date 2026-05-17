# SharePoint Search

`search` is a planned `sp-api` capability group. Inspect planned capabilities with:

```bash
sp-api schema
```

Planned verbs:

| Verb | Purpose |
|------|---------|
| `query` | Site-scoped SharePoint search |
| `people` | People-oriented search |

Until the group is implemented, report missing search coverage instead of falling back to raw HTTP. The intended implementation should expose query text, row limit, selected properties, and refiners as explicit schema params.
