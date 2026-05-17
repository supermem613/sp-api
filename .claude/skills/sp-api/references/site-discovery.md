# SharePoint Site Discovery

`sites` is a planned `sp-api` capability group. Basic list discovery is already implemented through `sp-api lists list`.

```bash
sp-api lists list
sp-api schema sites
```

Planned `sites` verbs:

| Verb | Purpose |
|------|---------|
| `get` | Get site metadata |
| `subsites` | List subsites |
| `navigation` | Read site navigation |

List schema discovery is currently available through:

```bash
sp-api lists get --title Tasks
sp-api schema lists get
```

Future field/content-type discovery should be added as semantic verbs rather than raw endpoint examples.
