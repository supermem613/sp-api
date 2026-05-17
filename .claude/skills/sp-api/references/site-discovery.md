# SharePoint Site Discovery

`sites` is implemented for current-site metadata and visible container discovery. Basic list discovery is also available through `sp-api lists list`.

```bash
sp-api sites get
sp-api sites discovery
sp-api lists list
sp-api lists fields --title Tasks
```

Implemented `sites` verbs:

| Verb | Purpose |
|------|---------|
| `get` | Get site metadata |
| `discovery` | List visible lists and libraries with root folder paths |

List field schema discovery is available through:

```bash
sp-api lists fields --title Tasks
sp-api schema lists fields
```

Prompt pattern:

> Inspect this site and tell me which lists, libraries, and fields are available before changing anything.

Use `sp-api sites get`, `sp-api sites discovery`, then `sp-api lists fields --title <list>` for the specific list. Future content-type, navigation, or subsites support should be added as semantic verbs rather than raw endpoint examples.
