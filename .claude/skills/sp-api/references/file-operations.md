# SharePoint File and Folder Operations

Use the `files` capability group for file work. The live command contract is generated from the registry:

```bash
sp-api schema files
sp-api files --help
```

## Implemented commands

| Task | Command |
|------|---------|
| List files in a folder | `sp-api files list --folder "/sites/team/Shared Documents"` |
| Get file metadata | `sp-api files get --path "/sites/team/Shared Documents/doc.txt"` |
| Download file content | `sp-api files download --path "/sites/team/Shared Documents/doc.txt"` |
| Upload small text content | `sp-api files upload --folder "/sites/team/Shared Documents" --name notes.txt --content "hello"` |
| Delete file | `sp-api files delete --path "/sites/team/Shared Documents/old.txt"` |
| Move file | `sp-api files move --path "/sites/team/Shared Documents/a.txt" --destination "/sites/team/Shared Documents/b.txt"` |
| Copy file | `sp-api files copy --path "/sites/team/Shared Documents/a.txt" --destination "/sites/team/Shared Documents/copy.txt"` |

## Schema-first workflow

```bash
sp-api schema files upload
sp-api schema files move
```

The schema describes required params, overwrite flags, underlying REST metadata, auth requirements, and response envelope shape.

## Paths

`files` commands use server-relative SharePoint paths, for example:

```text
/sites/team/Shared Documents/report.docx
```

Folder commands use the folder path. File commands use the full file path.

## Current limits

The current `files.upload` command is for small text content. Binary uploads, chunked uploads, folders, versions, check-in/check-out, and folder color should become explicit capability verbs before agents rely on them. Do not fall back to raw HTTP.
