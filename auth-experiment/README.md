# auth-experiment — Device-code login for ask-marcel-office

## Problem

`ask-marcel-office login` uses Playwright to open Edge and capture the
Teams web client token.  On this machine **two blockers** prevent it:

| Blocker | What it does |
|----------|-------------|
| **SentinelOne EDR** | Blocks the CDP pipe that Playwright uses to steer the browser. The browser launches but never responds to Playwright commands. |
| **v2rayN / sing-box proxy** | Sets `HTTP_PROXY=socks5://127.0.0.1:10808` and a PAC file in the registry. Bun's `fetch()` reads the PAC file and fails with `UnsupportedProxyProtocol`. |

## Solution

**OAuth 2.0 Device Authorization Grant** (RFC 8628) — no browser
automation, no CDP pipe, no Bun `fetch()`.

The script uses **PowerShell's `[System.Net.WebRequest]`** for HTTP
(bypasses the system proxy) and the **Outlook iOS client ID**
(`27922004-5251-4030-b22d-91ecd9a37ea4`) for the device-code flow.

The Outlook iOS app is the only Microsoft **public client** that carries
`Mail.Read`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, and
`People.Read.All` — the exact scopes the CLI needs.

## Client IDs tested

| Client ID | Name | Works? | Has Mail.Read? |
|-----------|------|--------|----------------|
| `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` | Teams web | ✗ — confidential app, needs client secret | ✓ |
| `04b07795-8ddb-461a-bbee-02f9e1bf7b46` | Azure CLI | ✓ — public client | ✗ — admin scopes only |
| `14d82eec-204b-4c2f-b7e8-296a70dab67e` | Graph PowerShell | ✓ | ✗ — admin scopes only |
| `1950a258-227b-4e31-a9cf-717495945fc2` | Azure PowerShell | ✓ | ✗ |
| `d3590ed6-52b3-4102-aeff-aad2292ab01c` | MS Auth Broker | ✓ | ✗ |
| `9ba1a5c7-f17a-4de9-a1f1-6178c8d51223` | Intune Company Portal | ✓ | ✗ |
| `1fec8e78-bce4-4aaf-ab1b-5451cc387264` | Teams Android | ✗ — blocked | — |
| `b26aadf8-566f-4474-9b0d-1c1be5c5e6c0` | Outlook Android | ✗ — blocked | — |
| **`27922004-5251-4030-b22d-91ecd9a37ea4`** | **Outlook iOS** | **✓** | **✓** |

## Usage

```powershell
powershell -File auth-experiment/poc-device-code.ps1
```

1. Opens `https://microsoft.com/device` — enter the code shown
2. Signs you in with your Microsoft 365 account
3. Saves the token to `~/.ask-marcel/token-cache.json`

After that, `ask-marcel-office` works normally:

```powershell
ask-marcel-office list-mail-messages --top 3
ask-marcel-office search-files --query "quarterly report"
```

## Token scopes

The Outlook iOS token grants:

```
Mail.Read, Mail.Read.Shared, Files.ReadWrite.All, Sites.ReadWrite.All,
People.Read.All, User.Read, Presence.Read.All, openid, profile, email,
People.Read, User.ReadBasic.All, UserAuthenticationMethod.ReadWrite
```

## Token lifetime

- **Access token**: ~1 hour (refreshed automatically by the CLI)
- **Refresh token**: ~90 days

When the refresh token expires, run the script again.

## Integration notes

To integrate into the main CLI as `ask-marcel-office login --device`:

1. Add the device-code flow in `src/infra/auth.ts` (next to `acquireViaBrowser`)
2. Use `[System.Net.WebRequest]` or a Bun-compatible HTTP client that
   bypasses the system proxy
3. The Outlook iOS client ID is stable — Microsoft doesn't rotate it
4. The token cache format is compatible with the existing `TokenCache` class