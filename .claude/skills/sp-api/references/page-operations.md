# SharePoint Page Operations

`pages` is implemented for listing, reading, field updates, checkout, publish, and discard-checkout for modern Site Pages.

```bash
sp-api schema pages
sp-api pages list
```

Implemented verbs:

| Verb | Purpose |
|------|---------|
| `list` | List site pages |
| `get` | Get page metadata/content |
| `checkout` | Check out a page by server-relative path |
| `save-fields` | Update Site Pages list item fields by item id |
| `publish` | Publish a page |
| `discard-checkout` | Discard a page checkout |

Prompt patterns:

> List pages and read the page named Home.

```bash
sp-api pages list
sp-api pages get --path "/sites/team/SitePages/Home.aspx"
```

> Update the page title and publish it.

```bash
sp-api pages checkout --path "/sites/team/SitePages/Home.aspx"
sp-api pages save-fields --item-id 42 --body '{"Title":"New title"}'
sp-api pages publish --path "/sites/team/SitePages/Home.aspx" --comment "Published by sp-api"
```

Use server-relative `--path` for file lifecycle verbs. Use `--item-id` for `save-fields` because it updates the backing Site Pages list item. Do not route agents to raw REST as a fallback.
