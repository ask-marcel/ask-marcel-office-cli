# 0002: Device-code login evaluated and not adopted

- Status: rejected for now (evaluated 2026-09-07); the findings are kept so the question is not re-derived
- Date: 2026-09-07

## Context

`login` drives a browser with Playwright and harvests the tokens Microsoft's own web clients mint (Teams web `5e3ce6c0`, M365ChatClient `c0ab8ce9`, OfficeHome `4765445b`). On one Windows machine this cannot work at all: a SentinelOne EDR policy blocks the local DevTools (CDP) pipe Playwright uses to steer the browser, so the browser opens, sits on `about:blank`, and is never driveable. No version of the CLI worked there under Bun, which framed the block as unfixable at the time (corrected below).

The OAuth 2.0 Device Authorization Grant (RFC 8628) needs no browser automation and no local pipe: the user enters a short code at `microsoft.com/device` on any device. A proof of concept was built in an `auth-experiment/` folder (since deleted; this record replaces it) to see whether device code could reach the same command surface without browser capture.

## What was learned

The CLI's 188 scoped commands need **25 distinct delegated scopes** (the union of `src/use-cases/commands/graph-scopes.ts`):

```
Calendars.Read Calendars.Read.Shared Channel.ReadBasic.All Chat.ReadBasic
ChatMember.Read Files.Read Files.Read.All Files.ReadWrite Group.Read.All
GroupMember.Read.All InformationProtectionPolicy.Read Mail.Read Mail.Read.Shared
Mail.ReadWrite MailboxSettings.Read Notes.Read Notes.Read.All People.Read
Place.Read.All Sites.Read.All Tasks.Read Team.ReadBasic.All
TeamsAppInstallation.ReadForTeam User.Read User.Read.All
```

The first POC used the Outlook iOS client with `/.default`, which returns only what a client is already consented for. That token covered **97 of 188** commands: files, SharePoint and Excel near-complete; mail 26 of 46; calendar, tasks, OneNote and Teams entirely absent. Two levers were then identified:

1. Requesting the 25 scopes **explicitly** against a public client that permits dynamic consent, instead of `/.default`. This is the only route to a single token with the full set, and it is what makes the mobile-app client hunt the wrong strategy: those clients are pre-consented for their own slice and drop scopes outside it at consent time.
2. Tenant policy, which is the real ceiling and the reason for this decision.

Which first-party clients even accept a device-code request was probed empirically against the `devicecode` endpoint, which rejects a blocked, confidential or unauthorized client without a sign-in:

| Client | ID | Result |
|---|---|---|
| Graph PowerShell | `14d82eec-204b-4c2f-b7e8-296a70dab67e` | accepts, dynamic consent, the parity candidate |
| Azure CLI | `04b07795-8ddb-461a-bbee-02f9e1bf7b46` | accepts, dynamic consent |
| Azure PowerShell | `1950a258-227b-4e31-a9cf-717495945fc2` | accepts |
| Outlook iOS | `27922004-5251-4030-b22d-91ecd9a37ea4` | accepts, mail/files slice under `.default` |
| Microsoft Office / broker | `d3590ed6-52b3-4102-aeff-aad2292ab01c` | accepts |
| OneDrive SyncEngine | `ab9b8c07-8f02-4f72-87fa-80105867a763` | accepts |
| Microsoft To-Do | `22098786-6e16-43cc-a27d-191a01a1e3b5` | accepts |
| Microsoft Edge | `ecd6b820-32c2-49b6-98a6-444530e5a77a` | accepts |
| Visual Studio | `872cd9fa-d31f-45e0-9eab-6e460a02d1f1` | accepts |
| Intune Company Portal | `9ba1a5c7-f17a-4de9-a1f1-6178c8d51223` | accepts |
| Office 365 Management | `00b41c95-dab0-4487-9791-b9d2c32c80f2` | accepts |
| Outlook Android | `b26aadf8-566f-4474-9b0d-1c1be5c5e6c0` | needs the tenant GUID as authority; `/common` and `/organizations` both AADSTS50059 |
| Teams desktop | `1fec8e78-bce4-4aaf-ab1b-5451cc387264` | refused, not enabled |
| Teams web, M365ChatClient, OfficeHome | the browser-capture identities | refused, "not supported for this feature": device code can never reuse them |
| `2d7f3606-b07d-41d1-b9d2-0d0c9296a6e8` | commonly mislabelled OneNote; it is Bing Search | refused |

Accepting the request is not the same as the tenant granting the scopes. That is decided at sign-in and the token's `scp` claim is the only truth. `/.default` also drags in scopes nobody wants cached for ninety days; the first POC's token carried `UserAuthenticationMethod.ReadWrite`, which can alter MFA methods.

## Decision

Keep browser capture as the only login. Do not ship `login --device`, for now.

The two methods are opposite trades. Browser capture gives a **fixed** scope set that works in every tenant, because Microsoft's first-party apps are pre-authorized globally and need no per-tenant consent; that is also why the set can never grow. Device code gives scopes you choose, gated per tenant by user-consent policy, admin-consent-required permissions (`Group.Read.All`, `User.Read.All`, `Sites.Read.All`, `Place.Read.All`, `Notes.Read.All` are the usual ones), Conditional Access (which can block the device-code flow outright or demand a managed device a CLI cannot be), and whether the signing user happens to be an admin.

A feature whose reach differs by tenant, and which cannot know its own reach until a token comes back, cannot be a default login. It could only ever be an additive tier with an honest per-run report, and the maintainer's stated requirement is a fixed, predictable scope set. That settles it.

## Consequences

- The EDR-blocked machine has a simpler fix than this ADR first assumed. **Correction 2026-09-08:** the SentinelOne block hit only under **Bun**; the same machine runs `login` fine under **Node**, and the published bin is `#!/usr/bin/env node`, so `npm i -g ask-marcel-office-cli` then `ask-marcel-office login` works. SentinelOne trusts the signed `node.exe` and blocks the less-common `bun.exe` from driving a child browser over CDP. So the block was never unfixable from the CLI: run under Node. A security exclusion, or copying `~/.ask-marcel/token-cache.json` from an unrestricted machine, remain fallbacks. The `Authentication cancelled` and elevated-timeout messages name both the cause and the run-under-Node workaround. This does not reopen the decision below: device code was rejected for the fixed-scope reason, which is independent of the EDR block that prompted the exploration.
- `Mail.Read.Shared` remains out of reach. The device grant carries it, and a real consent grant might read a shared mailbox where the harvested Outlook Web token was refused; that test was never run and stays open.
- Nothing in `src/` changed for this evaluation.

## If this is ever revisited

1. Request the 25 scopes explicitly against Graph PowerShell or Azure CLI, never `/.default`.
2. The refresh path in `src/infra/auth.ts` hardcodes the Teams client id when redeeming a refresh token (`redeemRefreshToken`, the basic-refresh callsite). A device-code refresh token can only be redeemed by its issuing client, so the cache must carry `client_id` and the refresh must use `cached.client_id ?? CLIENT_ID`, defaulting to the Teams client for existing caches. Without this the device token dies after roughly an hour.
3. The device login must merge into the existing cache the way `refreshToken` does, never replace it, or it wipes the elevated, chatsvcagg and ic3 tiers on a browser-capable machine.
4. Never touch machine proxy configuration; the POC cleared registry proxy keys and did not restore them.
5. Test on more than one tenant before generalising: decant's second tenant may grant a different set.
