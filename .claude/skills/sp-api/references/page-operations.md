# SharePoint Page Operations

`pages` is a planned `sp-api` capability group. Inspect planned capabilities with:

```bash
sp-api schema
```

Planned verbs:

| Verb | Purpose |
|------|---------|
| `list` | List site pages |
| `get` | Get page metadata/content |
| `create` | Create a page |
| `update` | Update page metadata/content |
| `publish` | Publish a page |
| `delete` | Delete a page |

Modern page APIs have SharePoint-specific edge cases such as publish/edit conflicts. Capture those as schema docs and command behavior when the capability is implemented. Do not route agents to raw REST as a fallback.
