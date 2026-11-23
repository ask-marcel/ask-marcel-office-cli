# The token tiers: facts for contributors

This doc is for contributors to the auth package. It is not in the npm package.
For the user view of the tiers, read [USAGE.md](USAGE.md).

Most facts here come from live probes against Microsoft 365. The date of each
probe is in parentheses. Microsoft can change a grant, a host or a page at any
time, so a fact is true at the date of its probe. Before you change auth code on
one of these facts, do the probe again.

## The scopes of each token

Use the basic tier first. Use the elevated tier only for an endpoint that gave a
403 to the basic token in a live probe, not for an assumption from the docs. The
basic token reads `GET /users/{id}` with `department` and `$expand=manager`
(2026-08-15), and `GET /chats/{id}/members` with `ChatMember.Read` (2026-06-23).
`/me/chats` and `/chats/{id}` need `Chat.ReadBasic`, which only the elevated
token has.

The scope set of the Teams web client token stays the same for a tenant only
until Microsoft changes the grants of the app. On 2026-09-09 the basic token had
`ChannelMessage.Read.All`, `TeamMember.Read.All` and `TeamsTab.Read.All` (as
`.ReadWrite.All` where Microsoft grants the write form). The snapshot of
2026-04-30 did not have them, and `/me/chats` stayed a 403. A scope in the token
is a prediction: the Outlook token had `Mail.ReadWrite.Shared`, and Graph
refused it. Read `status` again at each QA run (playbook E3), and probe the
endpoint before a command ships on a scope. A 403 is not always a scope failure:
an unknown channel message id gives `403 Forbidden: UnknownError` (2026-09-09).

Neither token has `Mail.Read.Shared` (2026-08-30):

- The basic token (client id `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`,
  `graph/.default`) has `Mail.Read`, `Mail.ReadWrite`, `Calendars.Read.Shared`
  and `Place.Read.Shared`.
- The elevated token has no shared mail scope.
- Two tenants gave the same 37 scopes, so the app registration sets the limit,
  not the consent of a tenant.
- Outlook on the web (app `9199bf20-a13f-4107-85dc-02114787ef48`) gives 7
  tokens. The token for `https://outlook.office.com` has
  `Mail.ReadWrite.Shared` and `Mail.Send.Shared`. The token for Graph has no
  mail scope.

We do not use the Outlook token: it can write and send, and the read package
must stay read-only. Another mailbox answers `ErrorAccessDenied`. That code does
not tell a missing scope from a missing Exchange delegation. A `scopesRequired`
entry predicts a 403. It is not a promise that the command works.

## The captures

The browser sign-in opens the Teams web client for the basic token. Then it
opens `https://m365.cloud.microsoft/search` (`M365_CLOUD_URL`) for the elevated
token. The capture opens `/search` and not the root. The root sometimes stops on
`/chat/blocked` or does not complete its start. `/search` makes the page ask for
a Graph token for OfficeHome and M365ChatClient in the first 5 to 8 seconds
after `domcontentloaded` (2026-05-06).

The elevated capture uses a visible browser window. `m365.cloud.microsoft` gave
a headless Edge an interstitial page and no token (2026-05-06). The cookies of
the browser profile do the sign-in, so the window closes with no input. The
short window at the expiry of the token is correct.

Two apps give an elevated Graph token on the same page load, in random order,
with gaps of less than one second (2026-08-30). The apps are M365ChatClient
(`c0ab8ce9-…`) and OfficeHome (`4765445b-…`), both in `ELEVATED_APP_IDS`:

- Only M365ChatClient has `Chat.ReadBasic`, `Sites.ReadWrite.All` and
  `Group.Read.All`.
- OfficeHome has `Contacts.Read` and not these three.
- The capture prefers M365ChatClient. It holds an OfficeHome token for
  `elevatedPreferenceGraceMs` (3 s), and it keeps that token only if
  M365ChatClient does not come.

Before you accept a second app for a tier, decode and compare the `scp` claims.

The caller sets the deadline of the elevated capture. `login` gives the long
deadline (`pollDeadlineMs`), because a person waits. A command gives the short
cap (`elevatedRecaptureTimeoutMs`, 20 s), so that a background command does not
hang. The code does not look at the page URL: a silent sign-in also goes through
`login.microsoftonline.com`. A deadline from the URL closed windows that were
about to succeed (2026-08-31).

A capture that succeeds took 41 s with the bundled binary and 104 s from source
(2026-08-31), and about 10 s after `71c8e6d`. After a completed sign-in, the
profile holds `ESTSAUTHPERSISTENT` with an expiry of 90 days (claim
`signin_state: ["kmsi"]`), and each later plain `login` keeps it. The elevated
token cannot come from the shared refresh token. A redemption for
M365ChatClient (`c0ab8ce9`) or OfficeHome (`4765445b`), with or without
`Origin: https://teams.microsoft.com`, answers `AADSTS70000 invalid_grant`
(2026-08-15). A token in the cache is no proof of a browser profile that is
signed in.

`login` tries the silent capture first (`getElevatedAccessToken()`, 15 to 17 s,
cookies untouched). It uses the forced sign-in on each failure, with no branch
by reason (2026-08-30). On the command path of the auth bin,
`recaptureElevatedViaBrowser` is true only at a terminal
(`auth/src/composition/build-auth-deps.ts`). Thus an MCP or CI run gets
`secondary_token_unavailable` immediately. The elevated tier has a flag of its
own: `chatsvcagg` and `ic3` renew from the refresh token with no browser, and a
shared flag replaced that renewal with a browser.

`login --force` clears the cookies (`clearCookies()` in `acquireBothTokens`).
The basic token comes from the response of the OAuth grant, so the grant must
fire again while the listener is attached. The clear also deletes
`ESTSAUTHPERSISTENT`, the cookie of about 90 days that Entra writes at "Stay
signed in?". Only session cookies (`AADSSO`, `esctx`) and telemetry cookies
(`buid`, `fpc`, `brcap`, `OH.FLID`) stay (2026-08-30). A clear of only
`localStorage` and `sessionStorage` is not proven. Before you change
`acquireBothTokens`, do a live probe:

1. Clear the two web storages in a persistent context.
2. Load `teams.microsoft.com` again.
3. Look for a `/oauth2/v2.0/token` response while the listener is attached.

To see if a profile has a session, read `is_persistent` and `expires_utc` of
`ESTSAUTHPERSISTENT` in `Default/Cookies` with sqlite3.

`teams.microsoft.com` redirects (302) to `teams.cloud.microsoft`, and each
substrate call then goes to the new host (2026-08-31). Both hosts serve the
substrate API in the same way. `parseChatsvcaggRegion` accepts the two hosts by
name and no wildcard, because the region goes into request URLs. With no
region, the capture used its full 20 s deadline and then the `emea` default.
That default gives wrong data with a success envelope on an APAC or AMER tenant.
When Microsoft changes a host, search the matchers for the old host, not only
the navigation targets.

## The lifetimes and the refresh

The refresh token lives about 24 hours from the browser sign-in
(`AADSTS700084`, 2026-11-19). [USAGE.md](USAGE.md) gives the details. For a live
run, start in the 24 hours after a browser sign-in. Do not start a sign-in that
you cannot complete.

The browser sign-in captures `chatsvcagg` and `ic3` only when Teams sends their
requests (2026-07-13). The chat list on the landing page gives `chatsvcagg`.
Only the load of a chat history gives `ic3` (`aud=https://ic3.teams.office.com`).
That load does not usually occur before the page goes to `m365.cloud.microsoft`.
A missing `ic3` after a plain `login` is not a defect. The reliable path is the
redemption from the refresh token (`refreshSubstrateToken`). The sign-in
redeems a missed tier after the capture (`redeemMissedSubstrateAtLogin`).

A second sign-in makes the server refuse the `chatsvcagg` and `ic3` tokens of
the session before, although their expiry is hours away (2026-08-31: issued at
07:18, bare 401 at 13:47, on `teams.microsoft.com` and on
`teams.cloud.microsoft`). The expiry tells when a token stops, not that it is
good. On a 401 the client drops the token, gets a new one from the refresh token
over HTTP and sends the request one more time (`substrateGet` in
`core/src/infra/graph-substrate.ts`). A remedy hint must name a command that can
repair that state.

## The guest tier

`401 invalidAudienceUri` with the app id of SharePoint Online
(`00000003-0000-0ff1-ce00-000000000000`) on a file of a partner tenant shows the
tenant border (2026-07-16). It is not a scope problem and not the ODSP
allow-list, and `Prefer: redeemSharingLink` does not help.

- The read package finds the tenant from the URL host. `<name>.sharepoint.com`,
  without `-my`, gives `<name>.onmicrosoft.com`. The public OIDC discovery
  document of that tenant gives the tenant id.
- Auth redeems the shared refresh token at
  `login.microsoftonline.com/{tid}/oauth2/v2.0/token` with
  `Origin: https://teams.microsoft.com`. Without the header, Entra answers
  `AADSTS9002327` before it looks at the grant, which looks like a refusal of
  the tenant.
- The guest token reads the full `/drives` surface of that tenant (metadata 200,
  `/content` 302, `?format=pdf` 302).
- One refresh token slot is correct: a guest redemption rotates the token, and
  the home tenant still renews.

The host rule is a convention. A failed discovery means "cannot cross to this
tenant", not a crash. This border is not the border of a directory lookup: the
home object id of a user in another tenant stays out of reach through `/users`.
