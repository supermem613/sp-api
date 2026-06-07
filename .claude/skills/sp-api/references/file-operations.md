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
| Expand child folders and files | `sp-api files folder --folder "/sites/team/Shared Documents"` |
| Get file metadata | `sp-api files get --path "/sites/team/Shared Documents/doc.txt"` |
| Download file content | `sp-api files download --path "/sites/team/Shared Documents/doc.txt"` |
| Download directly to a local file (binary-safe) | `sp-api files download --path "/sites/team/Shared Documents/image.png" --out image.png` |
| Upload small text content | `sp-api files upload --folder "/sites/team/Shared Documents" --name notes.txt --content "hello"` |
| Upload UTF-8 text from a local file | `sp-api files upload --folder "/sites/team/Shared Documents" --name SKILL.md --content-file SKILL.md` |
| Create folder | `sp-api files create-folder --path "/sites/team/Shared Documents/New Folder"` |
| Delete folder | `sp-api files delete-folder --path "/sites/team/Shared Documents/Old Folder" --missing-ok true` |
| Recycle folder | `sp-api files recycle-folder --path "/sites/team/Shared Documents/Old Folder" --missing-ok true` |
| Delete file | `sp-api files delete --path "/sites/team/Shared Documents/old.txt"` |
| Recycle file | `sp-api files recycle --path "/sites/team/Shared Documents/old.txt"` |
| Move file | `sp-api files move --path "/sites/team/Shared Documents/a.txt" --destination "/sites/team/Shared Documents/b.txt"` |
| Copy file | `sp-api files copy --path "/sites/team/Shared Documents/a.txt" --destination "/sites/team/Shared Documents/copy.txt"` |

## Schema-first workflow

```bash
sp-api schema files upload
sp-api schema files create-folder
sp-api schema files delete-folder
sp-api schema files recycle-folder
sp-api schema files folder
sp-api schema files recycle
sp-api schema files move
```

The schema describes required params, overwrite flags, `--content-file`, `--missing-ok`, underlying REST metadata, auth requirements, and response envelope shape.

For `files download`, use `--out <local-path>` to write the raw response bytes to disk instead of returning the content string in `data`.

## Paths

`files` commands use server-relative SharePoint paths, for example:

```text
/sites/team/Shared Documents/report.docx
```

Folder commands use the folder path. File commands use the full file path.

## Upload from local text files

Use `--content-file` when the body is easier to manage as a local UTF-8 file:

```bash
sp-api files upload --folder "/sites/team/Shared Documents/New Folder" --name SKILL.md --content-file SKILL.md
```

`--content` and `--content-file` are mutually equivalent sources for the upload body. One is required.

## Folder CRUD

Folder verbs are semantic commands. Do not fall back to raw REST:

```bash
sp-api files create-folder --path "/sites/team/Shared Documents/New Folder"
sp-api files delete-folder --path "/sites/team/Shared Documents/New Folder" --missing-ok true
sp-api files recycle-folder --path "/sites/team/Shared Documents/New Folder" --missing-ok true
```

`--missing-ok true` turns a missing folder into a successful no-op. Other SharePoint failures, including auth and permission failures, still fail loudly.

## Current limits

The current `files.upload` command is for small text content. Binary uploads, chunked uploads, versions, check-in/check-out, and folder color should become explicit capability verbs before agents rely on them. Do not fall back to raw HTTP.
