# SharePoint Search

`search` is an implemented `sp-api` capability group for bounded SharePoint search.

```bash
sp-api schema search
sp-api schema search query
```

Implemented verbs:

| Verb | Purpose |
|------|---------|
| `query` | Site-scoped SharePoint search |

Prompt pattern:

> Search this SharePoint site for `<terms>` and return the title and path of the best matches.

Use:

```bash
sp-api search query --query "<terms>" --row-limit 10 --select-properties Title,Path,LastModifiedTime
```

Do not ask agents to call raw `_api/search/query`. If people search or refiners are needed, inspect `sp-api schema` and report the missing semantic verb.
