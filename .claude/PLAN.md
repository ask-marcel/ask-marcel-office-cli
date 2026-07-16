# PLAN: cross-tenant guest-token support

Read a OneDrive / SharePoint file whose sharing URL belongs to a tenant the signed-in user
is a GUEST in. Today every such call dies at `401 invalidAudienceUri: Invalid audience Uri
'00000003-0000-0ff1-ce00-000000000000'` (SharePoint Online's app id): home-tenant Graph
cannot mint an SPO token for a foreign tenant's SharePoint.

## Established by live probe (2026-07-16), not assumed

| Fact | Evidence |
|---|---|
| It is NOT scopes | basic + elevated tokens carry identical `Files.ReadWrite.All` + `Sites.ReadWrite.All` |
| It is NOT the ODSP appid allow-list | elevated (M365ChatClient) fails identically to basic |
| It is NOT `Prefer: redeemSharingLink` | all 4 combos (basic/elevated x ±redeem) fail identically |
| The FOCI RT crosses tenants | redeemed against `/{tid}/oauth2/v2.0/token` + `Origin: https://teams.microsoft.com` -> `tid=<foreign>` |
| `Origin` is REQUIRED | without it: `AADSTS9002327` (SPA client-type may only redeem via cross-origin requests) |
| Tenant is derivable from the URL host | `{prefix}.sharepoint.com` -> `{prefix}.onmicrosoft.com` -> OIDC discovery -> tid |
| The guest token works on the WHOLE `/drives` path | metadata 200, `/content` 302, `?format=pdf` 302 via mediap.svc.ms |
| RTs are NOT tenant-scoped | after 2 foreign redemptions, `/common` still returns `tid=<home>` -> ONE shared RT slot is safe |
| `fetchUrl` allow-list already covers it | `graph-client.ts:155` is pattern-based (`/\.sharepoint\.com$/`, `/\.svc\.ms$/`) |

## Shape

`resolve-drive-share-link` returns `tenantId`; the drive-item family takes an OPTIONAL
`--tenant-id`. Optional flags leave the `required:boolean -> commander .requiredOption()`
wiring untouched, so `CommandMeta` needs no contract change. Rejected: one-shot
`download-share-link-*` commands (fail identically; and LESSONS 2026-07-06 says never ship a
wrapper for composable surface), `--share-url` XOR (no XOR precedent in 244 commands).

Guest is a RUNTIME choice, not a static one like `elevated` — the same command uses
`graph.get` without `--tenant-id` and `graph.getGuest` with it. So the builders ROUTE on the
parsed param; they do NOT grow guest twins (that would take 10 builders to 14+).

## Binding lessons (from .claude/LESSONS.md)

- **2026-07-15 [mistake]** every cache/state write MERGES (`{...existing, changed}`); a partial
  write silently drops sibling tokens. Assert the PERSISTED side effect, not just the return.
- **2026-07-13 [mistake]** a fake models your ASSUMPTION of an adapter, not the adapter. Test
  internal control flow at the adapter's OWN seam (`createAuthManagerFromApi` + `installFetchMock`).
- **2026-06-23 / 2026-07-13 x3 [mistake]** green fakes prove logic, never assumptions about the
  API. LIVE-SMOKE every auth-flow change before declaring done.
- **2026-06-15 [decision]** a new `GraphClient` method goes through `src/test-helpers/graph-client-fake.ts`.
- **2026-06-15 [gotcha]** every coverage tier gates at 100% here, not atelier's 80%.
- **2026-06-16 [gotcha]** "lint passes" means `bun run lint:strict`, not `bun run lint`.
- **2026-06-16 [decision]** trunk-based: commit to `main`, never branch, never PR.
- **2026-07-16 [decision]** cross-tenant DIRECTORY resolution (`/users/{home-id}`) is confirmed
  dead. This is a DIFFERENT boundary (token acquisition) — say so in the new entry or the next
  session will read them as contradictory.

## PRIVACY (rule 34 + the 2026-07 history purge)

Real tenant strings must NEVER reach tests, docs, fixtures, or commit messages: no `contoso`,
`contoso`, no tenant GUIDs, no `an internal project`. Use Contoso / Fabrikam.
Repo history was rewritten once to purge exactly this class of string.

## Slices (each <=10 files / <=300 lines, green at every commit)

1. [x] **domain: `TenantId`** — `src/domain/tenant-id.ts` + `tenant-id.test.ts`, 5 tests.
       Branded, GUID-shape factory returning `Result<TenantId, TenantIdError>`, mirroring
       `env-var.ts` / `access-token.ts` (flat in `src/domain/`, NOT a `values/` subdir — that is
       not this repo's layout). Case is deliberately NOT normalized: Entra accepts either casing
       on the authority AND the value doubles as the per-tenant cache key, so rewriting it would
       split one tenant's cache across two entries.
       Rule 12 earns it: the value is interpolated into an auth URL whose POST body carries the
       refresh token; the GUID gate closes path traversal at the boundary.
       `spoHostToTenantDomain` MOVED to slice 4 — this repo does not directly test
       `src/domain/utilities/**` (`archive-status.ts`, `site-url-classifier.ts` have no test
       file; they are exercised through the command port, per rule 14). Shipping it here would
       leave it at 0% and fail the 100% domain gate.
       DoD: tests 4540/0 green, typecheck clean, `lint:strict` clean, coverage gate exit 0
       (`tenant-id.ts` 100.00/100.00). Mutation: PENDING.

2. [x] **infra: guest-token rung in `auth.ts`** — `AuthManager.getGuestAccessToken(tenantId)`,
       9 tests. Cache field `guest_tokens?: Partial<Record<string, {access_token, expires_on}>>`,
       read via `Object.hasOwn`, MERGE-written after a RE-READ (the awaited network call means
       the on-disk cache may have moved; merging into the pre-await snapshot IS the 2026-07-15
       clobber). Rotated RT persisted. Rung logged as `auth.ladder.rung {rung:'guest'}`.
       Rule of Three honoured: `refreshToken` + `refreshSubstrateToken` + guest all collapse onto
       one `redeemRefreshToken(refreshToken, authority, scope)`. That core now carries
       `AbortSignal.timeout(REQUEST_TIMEOUT_MS)` — `auth.ts` previously set NO deadline on any
       token fetch (pre-existing rule 29 gap); once the three became one fetch a per-caller
       exception was arbitrary. FLAG THIS IN THE COMMIT — it changes shipped behaviour.
       Tenant DISCOVERY is NOT here — it is a plain unauthenticated OIDC fetch, not token
       acquisition. It lands in slice 3 as `graph.discoverTenantId(spoHost)` so the use-case
       orchestrates the retry.
       SIDE QUEST (approved): `AuthManager` was faked by 29 inline literals across 5 files, so
       adding one method broke all of them — the exact shape of the 2026-06-15 lesson, which
       says it covers "a similarly widely-faked port". Built `src/test-helpers/auth-manager-fake.ts`
       and migrated all 29. The next `AuthManager` method touches ONE file.
       DoD: tests 4548/0, typecheck clean, `lint:strict` exit 0, coverage `auth.ts` 100.00/100.00
       gate exit 0. Mutation N/A (scope is `^src/(domain|use-cases)/`; infra is out).

3. [ ] **infra: `graph.getGuest` / `graph.getBinaryGuest` + shared fake**
       Mirror the existing `getElevated` / `getBinaryElevated` pair. Update
       `src/test-helpers/graph-client-fake.ts` in the SAME commit.
       DoD: all existing tests green (the fake builder is why this stays a 1-file fake change).

4. [ ] **use-case: `resolve-drive-share-link` returns `tenantId`**
       Also lands `src/domain/utilities/spo-tenant.ts` (`spoHostToTenantDomain`, pure, null for
       non-SPO hosts incl. `1drv.ms`), exercised through this command's port rather than a
       direct test.
       DESIGN (refined after reading the code): the command does NOT detect "foreign" up front.
       It has no home-tenant id to compare against — `getCachedTokenInfo()` exposes no `tid` —
       so instead: try the home token, and ONLY on `401 invalidAudienceUri` derive the host's
       tenant domain, discover the tid, and retry as guest. The error IS the signal. Zero cost
       on the home path, no home-tid plumbing, self-healing. Explicit `--tenant-id` still exists
       for the drive-item commands, which hold only a `driveId` and cannot recover on their own.
       DoD: LIVE-verified against a real foreign share link (lesson 2026-06-23).

5. [ ] **builders route on `tenantId`**
       `tenantIdSchema` (optional) + `tenantIdOption` meta mirroring `odataQueryOptions`; the 6
       non-elevated builders route via a shared helper. Elevated builders do NOT get it
       (elevated is a home-tenant ODSP identity; cross-tenant elevated is not a thing).
       BLOCKER: adding `tenantId` to the meta.test.ts runtime-additive exclusion list EDITS AN
       EXISTING TEST -> rule 24, needs explicit user sign-off first.
       DoD: `meta.test.ts` invariants green; builder-based drive commands accept `--tenant-id`.

6. [ ] **hand-written executes thread `tenantId`**
       `FetchOptions` in `fetch-raw-bytes.ts` gains `tenantId?`; `inlineBinary` / `fetchRawBytes` /
       `officeToMarkdown` route. Covers download-drive-item-{content,as-markdown,as-pdf},
       extract-drive-item-images, convert-drive-item-zip, get-drive-item.
       DoD: LIVE end-to-end — the real foreign pptx converts via `--tenant-id`.

7. [ ] **docs + CHANGELOG + LESSONS**
       Regen `docs/commands.json` + `COMMANDS.md`; CHANGELOG entry; README audit (guideline #5);
       propose LESSONS entries.
       DoD: generated docs match source; README audited; entries proposed, not appended.

## Found in-flight (candidate LESSONS entries, propose at wrap-up)

- **[gotcha] `mutate:changed` cannot see an untracked file — a NEW file silently skips the
  mutation gate.** `scripts/mutate-changed.sh` selects via `git diff --name-only <BASE>...HEAD`
  + `git diff HEAD` + `git diff --cached`. A brand-new, never-`git add`-ed file appears in NONE
  of those, so `src/domain/tenant-id.ts` was never mutated while the script exited 0 and looked
  like a pass. Only `git add` (making it `--cached`) or an explicit
  `bunx stryker run --mutate <path>` reaches it. Every new domain/use-case file this repo has
  ever added hit this. Rule: for a NEW file, stage it first or mutate it by explicit path;
  never trust `mutate:changed` alone on a file git has not seen.
- **[gotcha] `mutate:changed` bases on `origin/main`, so unpushed commits widen the scope
  silently.** With local `main` 3 commits ahead of `origin/main`, it selected `get-user.ts`,
  `graph-scopes.ts`, and `resolve-drive-share-link.ts` — all already committed, none touched
  this session — and spent 29 min mutating them. Use `BASE=HEAD` (the script supports a `BASE`
  env override) when iterating with unpushed commits, or push first.

## Open questions

- Tenant discovery is unproven beyond n=1. A vanity SPO domain may not map to
  `{prefix}.onmicrosoft.com`. Fallback candidate: SPO's `WWW-Authenticate` realm on a 401.
  Decide in slice 2; do not silently assume the mapping holds.
- Consent: the Teams client may not be consented in an arbitrary partner tenant ->
  `AADSTS65001` / `interaction_required`. Needs a clean actionable error, not a timeout.
- CI cannot test cross-tenant (needs a live foreign tenant). Live-probe-only, like the
  substrate commands. Consider a QA-playbook row.
- Pre-existing, NOT introduced by this work: `writeCache` is a bare `fs.writeText` (no
  temp+rename) and there is no locking anywhere, so concurrent commands already race the
  single-use RT. Flag; do not fix in a feature commit (LESSONS 2026-07-15: a scoped change is
  not the venue to harden an unrelated pre-existing branch).
