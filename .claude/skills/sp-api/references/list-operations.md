# SharePoint List Operations

Use the `lists` capability group for list and list-item work. The live command contract is generated from the registry:

```bash
sp-api schema lists
sp-api lists --help
```

## Implemented commands

| Task | Command |
|------|---------|
| List visible lists | `sp-api lists list` |
| Get list metadata | `sp-api lists get --title Tasks` |
| Get field schema | `sp-api lists fields --title Tasks` |
| Create a list | `sp-api lists create --title "Project Tracker"` |
| Delete a list | `sp-api lists delete --title "Project Tracker"` |
| List items | `sp-api lists items --title Tasks --select Title,Id,Status --top 25` |
| Add item | `sp-api lists add-item --title Tasks --body '{"Title":"New task"}'` |
| Update item | `sp-api lists update-item --title Tasks --item-id 42 --body '{"Status":"Done"}'` |
| Delete item | `sp-api lists delete-item --title Tasks --item-id 42` |

## Schema-first workflow

Before calling a command, inspect its generated schema:

```bash
sp-api schema lists add-item
sp-api schema lists fields
sp-api schema lists update-item
```

The schema includes required params, underlying REST metadata, auth requirements, examples, and the JSON envelope shape.

## Metadata type for item creation

Some SharePoint lists require `__metadata.type` in item bodies. Get it with:

```bash
sp-api lists get --title Tasks
```

Use the returned `ListItemEntityTypeFullName` value in mutation bodies when SharePoint requires it:

```bash
sp-api lists add-item --title Tasks --body '{"__metadata":{"type":"SP.Data.TasksListItem"},"Title":"New task"}'
```

For field-level schema before creating or updating items, use:

```bash
sp-api lists fields --title Tasks
```

## Filtering and paging

The current `lists.items` command supports `--select` and `--top`. More complex OData filtering, CAML, field management, and views should become semantic verbs before agents rely on them. Do not fall back to raw HTTP.

Planned list verbs include field management, views, CAML query, batch updates, item versions, and recycle/restore.
