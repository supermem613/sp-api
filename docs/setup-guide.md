# sp-api Setup Guide

This guide covers installing and authenticating the `sp-api` CLI. The repo also ships a thin skill router that agents can load after the CLI is installed.

---

## Install

```bash
cd <repo-root>
npm install
npm run build
npm link
```

Verify the CLI:

```bash
sp-api doctor
sp-api schema auth login
```

`npm install` installs Playwright, which is used only by `sp-api auth`. `npm run build` validates the CLI, generated skill router, and bin wiring before `npm link` exposes `sp-api` on your PATH.

Update a linked or git-clone install later with:

```bash
sp-api update
```

## Authenticate

```bash
sp-api auth login --site contoso.sharepoint.com/sites/mysite
```

Replace `contoso.sharepoint.com/sites/mysite` with your site URL.

On first run, Edge may open a visible browser window. Sign in with your Microsoft account. Once login completes, the browser closes and your session is saved to a local profile.

Subsequent runs use the saved profile headlessly, extract cookies, and update `~/.sp-api/auth.json`.

## What Gets Saved

Auth credentials are saved to `~/.sp-api/auth.json`:

| Field | Description |
|-------|-------------|
| `SP_SITE` | Full site URL |
| `SP_COOKIES` | FedAuth and rtFa cookies for SharePoint REST calls |
| `SP_TOKEN` | Optional SharePoint bearer token captured from the browser session |

The auth file is an implementation detail consumed by `sp-api` and its internal helpers. Agents should not read or write it directly.

## Verify

```bash
sp-api auth status
sp-api lists list
```

You should receive JSON envelopes on stdout. For command details, run:

```bash
sp-api lists --help
sp-api schema lists list
```

## Login / Logout

Force visible re-login:

```bash
sp-api auth login --site contoso.sharepoint.com/sites/mysite --force
```

Clear the saved profile and auth file:

```bash
sp-api auth logout
```

The browser profile is stored at `~/.sp-api/browser-profile/`.

## Troubleshooting

### Edge not found

Playwright requires Microsoft Edge. Install Edge from [microsoft.com/edge](https://www.microsoft.com/edge).

### Login loop

Your saved session may have expired. Force a fresh login:

```bash
sp-api auth login --site contoso.sharepoint.com/sites/mysite --force
```

### No cookies found for tenant

1. Verify the site URL includes the SharePoint host and path.
2. Use `--force` for interactive login.
3. For dogfood tenants, use the full hostname such as `contoso.sharepoint-df.com`.

### HTTP 401

Cookies expired. Re-run `sp-api auth login --site <site>`.

### HTTP 403

Your browser account does not have permission to the requested SharePoint resource.
