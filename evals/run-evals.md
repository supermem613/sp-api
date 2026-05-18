# sp-api — sp-api Evals

Evals verify that agents use `sp-api`, pass arguments that match `sp-api schema`, and interpret JSON envelopes.

## How to Run

```text
Run evals/run-evals.md against <site-url>
```

## Execution Model

Agents must invoke `sp-api` commands. Do not use raw HTTP or direct reads of `~/.sp-api/auth.json`.

Each non-help command writes a JSON envelope to stdout:

```json
{
  "ok": true,
  "command": "lists.items",
  "data": {},
  "error": null,
  "meta": {
    "schemaVersion": "0.1.0"
  }
}
```

Pass conditions should inspect `ok`, `command`, and `data`.

## Setup

1. Authenticate:

   ```bash
   sp-api auth login --site <site-url>
   ```

2. Verify setup:

   ```bash
   sp-api doctor
   sp-api schema
   ```

3. Use a dev/test SharePoint site only. Evals create data prefixed with `SP_API_EVAL_`.

## Current v0.1 Evals

### 01 — Auth status

**Run:** `sp-api auth status`

**Pass if:** `ok` is `true`, `command` is `auth.status`, and `data.exists` is a boolean.

### 02 — Full schema

**Run:** `sp-api schema`

**Pass if:** `ok` is `true` and `data.capabilities.lists`, `data.capabilities.files`, and `data.plannedCapabilities` exist.

### 03 — Focused list schema

**Run:** `sp-api schema lists add-item`

**Pass if:** `ok` is `true`, `data.id` is `lists.add-item`, and `data.params` includes `title` and `body`.

### 04 — List visible lists

**Run:** `sp-api lists list`

**Pass if:** `ok` is `true`, `command` is `lists.list`, and `data.value` is an array.

### 05 — Create eval list

**Run:** `sp-api lists create --title SP_API_EVAL_List`

**Pass if:** `ok` is `true` and response data contains an id. Save the list title.

### 06 — Get eval list

**Run:** `sp-api lists get --title SP_API_EVAL_List`

**Pass if:** `ok` is `true` and response data includes the eval list title.

### 07 — Add list item

**Run:** `sp-api lists add-item --title SP_API_EVAL_List --body '{"Title":"SP_API_EVAL_ITEM"}'`

**Pass if:** `ok` is `true` and response data contains the item title or id. Save the item id.

### 08 — List items

**Run:** `sp-api lists items --title SP_API_EVAL_List --select Title,Id --top 5`

**Pass if:** `ok` is `true` and `data.value` is an array.

### 09 — Update item

**Run:** `sp-api lists update-item --title SP_API_EVAL_List --item-id <ITEM_ID> --body '{"Title":"SP_API_EVAL_UPDATED"}'`

**Pass if:** `ok` is `true`.

### 10 — Delete item

**Run:** `sp-api lists delete-item --title SP_API_EVAL_List --item-id <ITEM_ID>`

**Pass if:** `ok` is `true`.

### 11 — File schema

**Run:** `sp-api schema files upload`

**Pass if:** `ok` is `true`, `data.id` is `files.upload`, and params include `folder`, `name`, `content`, and `content-file`.

### 12 — List files

**Run:** `sp-api files list --folder "<DOCLIB_PATH>"`

**Pass if:** `ok` is `true` and `data.value` is an array.

### 13 — Upload file

**Run:** `sp-api files upload --folder "<DOCLIB_PATH>" --name SP_API_EVAL.txt --content "Hello from eval"`

**Pass if:** `ok` is `true`.

### 14 — Download file

**Run:** `sp-api files download --path "<DOCLIB_PATH>/SP_API_EVAL.txt"`

**Pass if:** `ok` is `true` and `data` contains `Hello from eval`.

### 15 — Delete file

**Run:** `sp-api files delete --path "<DOCLIB_PATH>/SP_API_EVAL.txt"`

**Pass if:** `ok` is `true`.

### 16 — Create folder

**Run:** `sp-api files create-folder --path "<DOCLIB_PATH>/SP_API_EVAL_Folder"`

**Pass if:** `ok` is `true`.

### 17 — Recycle folder

**Run:** `sp-api files recycle-folder --path "<DOCLIB_PATH>/SP_API_EVAL_Folder" --missing-ok true`

**Pass if:** `ok` is `true`.

### 18 — Delete folder with missing-ok

**Run:** `sp-api files delete-folder --path "<DOCLIB_PATH>/SP_API_EVAL_Folder" --missing-ok true`

**Pass if:** `ok` is `true` even if the previous recycle removed the folder.

### 19 — Delete eval list

**Run:** `sp-api lists delete --title SP_API_EVAL_List`

**Pass if:** `ok` is `true`.

## Report

Write `evals/results/report.md`:

```markdown
# Eval Report — [date]

**Site:** [site URL]
**Overall:** [passed]/16 ([percentage]%)

| Eval | Result | Notes |
|------|--------|-------|
| 01 Auth status | PASS/FAIL | |
```
