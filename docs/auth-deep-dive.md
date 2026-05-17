# Authentication Deep Dive

A technical deep-dive into how `sp-api auth` uses Playwright persistent context authentication for SharePoint API access.

## How Playwright Persistent Context Auth Works

1. `sp-api auth login` invokes the internal Playwright auth helper
2. Playwright launches a Chromium-based browser (Microsoft Edge) using `chromium.launchPersistentContext`
3. The persistent context stores its profile at `~/.sp-api/browser-profile/`
4. On first run, the browser opens visibly and the user signs into SharePoint
5. The browser profile saves cookies, localStorage, and session state to disk
6. On subsequent runs, the browser launches headlessly and loads the saved profile
7. Cookies are extracted via `context.cookies()` and saved in `~/.sp-api/auth.json`
8. Semantic `sp-api` commands use the saved auth through internal REST helpers

### Why Persistent Context Matters

Unlike regular Playwright contexts that start fresh each time, a persistent context stores its state in a user data directory — just like a normal browser profile. This means:

- **Login persists** — SSO tokens, session cookies, and auth state survive across runs
- **Windows SSO/WAM integration** — if you're logged into Windows with your corp account, SharePoint auth may work automatically (no manual login at all)
- **No app registration** — the browser session has the same permissions as your normal browser
- **No secrets to manage** — no client IDs, client secrets, or certificates

### How Windows SSO/WAM Provides Frictionless Auth

On Windows machines joined to Azure AD (Entra ID), the Web Account Manager (WAM) provides Single Sign-On:

1. When you sign into Windows with your corporate account, WAM caches your auth tokens
2. Edge (built on Chromium) has native WAM integration
3. When Playwright launches Edge with `channel: 'msedge'`, it uses the real Edge binary (not Chromium)
4. Edge's WAM integration picks up your Windows login automatically
5. SharePoint recognizes the WAM-issued tokens and sets session cookies without a manual login prompt

This means on a corp-joined Windows machine, the very first run may complete without any manual login.

## Key SharePoint Cookies

- **FedAuth** — The main authentication cookie. Base64-encoded SAML token. Typically 1000–3000 chars.
- **rtFa** — "Remember the FedAuth" — a refresh/session cookie that works alongside FedAuth.
- Together, these two cookies provide full authenticated access to the SharePoint REST API.

## Cookie Properties

| Aspect | Session Cookie (Playwright) |
|--------|----------------------------|
| Format | Opaque (FedAuth: base64 XML, rtFa: opaque) |
| Lifetime | ~8–24 hours |
| Refresh | Re-run `sp-api auth login --site <site>` |
| Scope | Full (same as browser session) |
| Used for | SP REST API (`/_api/...`) |

## Profile Storage

The browser profile is stored at:

```
~/.sp-api/browser-profile/
```

This directory contains:

- Cookie database
- localStorage / sessionStorage data
- Cached auth tokens
- Browser state (history, preferences)

### Profile Lifecycle

| Event | What Happens |
|-------|-------------|
| First run | Profile created, Edge opens for login |
| Subsequent runs | Profile reused, headless, cookies extracted instantly |
| `sp-api auth login --site <site> --force` | Profile reused but Edge opens visibly for re-login |
| `sp-api auth logout` | Profile directory deleted entirely |
| Cookies expire | Auth detects login redirect, falls back to visible login |

## Security Considerations

### Profile on Disk

- The browser profile is stored in your home directory with standard file permissions
- It contains the same data as your normal Edge profile — cookies, cached auth tokens, etc.
- Anyone with access to your home directory can read these cookies
- This is the same security model as using Edge normally
- The profile does NOT contain your password — only session tokens

### Best Practices

- Don't log or commit cookie values
- Re-authenticate periodically (cookies expire)
- Use `sp-api auth logout` to clear the profile when switching accounts or machines
- The profile directory is excluded from version control by default (it's in your home directory, not the repo)

### Comparison to Previous CDP Approach

| Aspect | Playwright Persistent Context | CDP (previous) |
|--------|-------------------------------|----------------|
| Setup | `npm install` (one-time) | Close all Edge windows, relaunch with debug flag |
| Auth flow | Automatic (profile persists) | Manual (visit site in debug browser every time) |
| Dependencies | Node.js + Playwright | Python 3 + websocket-client (or PowerShell WebSocket) |
| Security | Profile stored at ~/.sp-api/ | Cookies only in env vars (lost on shell close) |
| Reliability | Very high (Playwright manages browser lifecycle) | Fragile (debug port conflicts, silent flag ignore) |
| Login persistence | Yes (profile survives across sessions) | No (must re-login every time Edge restarts) |
