# Lessons (committed)

Append-only institutional memory for this codebase. See the atelier skill's `references/lessons.md` for the format and rules.

Each entry is one of `[mistake]`, `[decision]`, or `[gotcha]`. Newest first.

---

## [gotcha] 2026-06-13 | an additive error-envelope field must be threaded through ALL THREE layers (infra → composition → presenter) or CLI consumers never see it

Surfacing Graph's `Retry-After` as `retryAfterSeconds` required edits in three files, not one: the `api_error` variant of `GraphError` (`src/infra/graph-client.ts` — `parseRetryAfter` reads `res.headers.get('retry-after')`, delta-seconds only, spread into all THREE `apiErrorFrom` return paths), the composition `fail()` wrapper + the per-command error site (`src/composition/cli.ts` — extract `result.error.type === 'api_error' ? result.error.retryAfterSeconds : undefined` and forward it), AND the presenter `renderError` (`src/presenter/output.ts` — add to the JSON envelope and a `retryAfter: Ns` text line). Adding the field to `GraphError` alone reaches *library* consumers (`commands[x].execute` callers) but NOT CLI consumers, because `renderError` takes discrete primitive params, not the whole error object — the same shape the `errorCode` / `hint` / `source` fields already follow. Guard on `!== undefined`, never truthiness: `0` is a valid "retry immediately" hint. The RFC HTTP-date form is intentionally unparsed (needs a clock; Graph uses integer seconds for throttling) — absent/non-numeric → field omitted → caller falls back to its own backoff. ask-marcel itself still does NOT retry; the client is fail-fast and the orchestrator owns the backoff loop.
Applies to: any new field on the `{ok, error, errorCode?, hint?, source}` envelope — thread GraphError → cli.ts `fail` → output.ts `renderError`, and document it in docs/USAGE.md (the canonical envelope-shape block) + README.

## [decision] 2026-06-13 | local-only (non-Graph) commands follow the `executeLocal(fs)` + redirect-`execute` pair; register them in COMMANDS_WITHOUT_SCOPES

`extract-local-file-images` is the SECOND local-only command (after `convert-local-file`), sharing the exact pattern: the registry-typed `execute(graph, params)` returns a 400 redirect error pointing at `executeLocal`, and the real entry point is the optional `executeLocal(fs, params)` on the `Command` type, which `cli.ts` routes to at the single execute call site (`cmd.executeLocal !== undefined ? executeLocal(fs, …) : execute(graph, …)`). It reuses `extractImagesFromBytes` (the same dispatch `extract-drive-item-images` uses — pure bytes-in/media-out, zero Graph dependency), so the new command is ~40 lines + a `FileSystem.readBytes` call. Every such command MUST be added to `COMMANDS_WITHOUT_SCOPES` in `graph-scopes.test.ts` (the central-scope-map invariant fails otherwise), use `category: 'meta'`, and a `graphPathTemplate` of the `(local) reads {path} …` form (so the meta invariant's "every option appears in the template" check passes). It completes the legacy-`.ppt` pipeline: `download-drive-item-as-pdf --output-path` → `convert-local-file` (text layer) → `extract-local-file-images` (slide images for OCR).
Applies to: any future command whose input is the local filesystem rather than Graph — mirror `convert-local-file` / `extract-local-file-images`, don't invent a new shape.

## [gotcha] 2026-06-13 | `^`-ranged lint toolchain drift can break `bun run lint` on an unchanged HEAD — fix rule severity at project level, never inline

`typescript-eslint` resolving to 8.59 (within the `^8.33.1` range in package.json) began flagging the `_graph` / `_params` parameters on local-only commands' redirect-`execute` (`@typescript-eslint/no-unused-vars` has NO default `argsIgnorePattern`), so lint failed with 2 errors on code that predates the session — a toolchain bump, not a code change. Fix is project-level in `eslint.config.js`: `'@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }]` with a comment naming the deliberate "required by the contract, unused by this implementation" marker. Per atelier rule 15 this is the sanctioned route (change severity/config at project level); an inline `// eslint-disable` is forbidden. Separately, ~29 prettier/prettier warnings across committed files came from the same kind of drift (eslint-plugin-prettier 5.5 re-wrapping long single-line object literals) — auto-fixable with `bunx eslint --fix`, spun off as its own style-only task rather than mixed into a feature commit.
Applies to: any "lint suddenly fails and I didn't touch that file" moment — suspect a `^`-range devDependency bump; fix at the rule/config level, never with an inline ignore.

## [gotcha] 2026-06-02 | the Search index returns non-navigable site TYPES (404s) the archive probe misses — filter by URL shape + a 404 probe

Beyond archived sites, `POST /search/query entityTypes:['site']` returns site resources whose `webUrl` 404s and which are NOT archived (so the 423 probe never catches them): SharePoint Add-in (app) webs on isolated `<tenant>-<hex>.sharepoint.com` hosts (they redirect to `/_layouts/15/AddinDeprecationAnnoucement.aspx`), SharePoint Embedded containers on `/contentstorage/...`, and `/_layouts/` system pages. `src/domain/utilities/site-url-classifier.ts` (`isNonNavigableSiteUrl`) drops these three by URL shape BEFORE probing (no Graph call → `nonNavigableExcluded`), and `filterOutArchivedSites` now also drops a probe that returns HTTP 404 (`notFoundExcluded`) — which is what finally removes an inaccessible personal OneDrive whose webUrl 404s but is not archived. Active OneDrives that probe OK are still kept. The host regex needs `-[0-9a-f]{10,}` (≥10 hex) so it does not match `-my`/`-admin`. Verified live 2026-06-02: 80 kept; 58 archived + 1 nonNavigable + 2 notFound excluded; a legit `…/sites/KRONOSTACON/…` on the normal host stayed while the add-in `-<hex>…/KRONOSTACON/_layouts/…` variant dropped.
Applies to: any "list the sites I can reach" surface; extend `site-url-classifier.ts` / the probe verdicts rather than special-casing in the command.

## [mistake] 2026-06-02 | `/tmp/...${Date.now()}.json` temp paths collide across parallel `bun test` processes → flaky

`build-deps.test.ts` seeded a real cache at `/tmp/atelier-build-deps-cache-${Date.now()}.json` and deleted it in `finally`. `Date.now()` (ms) is NOT unique across processes, so when several `bun test` runs overlap (CI shards, a `bun run` racing a background mutation run, or Stryker's own spawned dry-run) two land on the same path in the same millisecond and one's cleanup deletes the other's file mid-read → intermittent single-test failure. It surfaced only under load (0/6 idle runs, 1/4 parallel runs) and BLOCKED Stryker, whose initial dry-run must be 100% green. Fix: `mkdtempSync(join(tmpdir(), 'prefix-'))` for an OS-guaranteed-unique dir (node:fs is allowed in tests per hard rule 20; three infra tests already do this), write inside it, `rmSync(dir, {recursive:true})` in `finally`. To hunt a load-sensitive flake, run N suites in parallel rather than N sequentially.
Applies to: every test that touches a real temp path; never key a shared-filesystem path on `Date.now()` alone.

## [gotcha] 2026-06-02 | Stryker counts `Timeout` mutants as KILLED — manual `mutation.json` parsing must too

A mutant that makes a loop run forever (e.g. `start < n` → `start >= n`, or `start += chunk` → `start -= chunk`) is reported with `status: "Timeout"`, and Stryker's break-threshold math counts Timeout as detected/**killed**. A hand-rolled `mutation.json` parser that only counts `status === "Killed"` (treating Timeout as a survivor) under-reports the score and disagrees with Stryker's own pass/fail — e.g. a real 90.91% read as 89.3%. Trust the `Final mutation score of N` line in Stryker's log; if parsing the JSON, count `Killed` + `Timeout` (and exclude `CompileError`/`NoCoverage` per your gate's intent).
Applies to: any script that parses `reports/mutation/mutation.json` to gate or report mutation score.

## [decision] 2026-06-01 | site-search commands exclude archived sites via a per-site metadata probe (keep active OneDrives)

`search-all-accessible-sites` and `search-sharepoint-sites-by-name` now probe each returned site (`GET /sites/{id}?$select=id,webUrl,siteCollection`) and drop the archived ones, surfacing `archivedExcluded` / `archiveProbeErrors` (+ `archiveProbeTruncated` on search-all, capped at 250 probed sites to bound the fan-out and avoid 429). Detection is split atelier-style: pure predicate in `src/domain/utilities/archive-status.ts` (`isArchivedSite` reads `siteCollection.archivalDetails.archiveStatus` ∈ {recentlyArchived, fullyArchived, reactivating}; `errorIndicatesArchived` classifies the probe's error), the chunked/capped probe orchestration in `src/use-cases/commands/filter-archived-sites.ts`. Deliberate design (Vincent's call): keep ACTIVE OneDrives, exclude ONLY detected-archived — do NOT blanket-drop personal OneDrives by URL heuristic. A site whose probe fails for a non-archive reason is KEPT and counted in `archiveProbeErrors`, never silently lost. Live tenant run: 83 sites returned, 58 archived excluded, 3 probe errors — the user's own and colleagues' active OneDrives retained, the departed-employee archived OneDrive (`m_lopez_us_celine_com`) gone.
Applies to: the two site-search commands; any future "exclude unreachable/archived" filtering should extend the shared `archive-status.ts` predicate rather than add URL heuristics.

## [gotcha] 2026-06-01 | auto-archived OneDrives fail `GET /sites/{id}` with 423 resourceLocked, not 200 + archiveStatus

A departed/unlicensed user's OneDrive (the OneDrive service auto-archives sites unlicensed ≥93 days; browsing one redirects to `/_layouts/15/sharepointerror.aspx?scenario=SiteArchived`) does NOT return its archive state on the metadata resource. The Graph call ITSELF fails — `423 resourceLocked` ("Access to this site has been blocked") — the same signal `list-accessible-drives.ts` `isUnreachable` already maps to an admin-locked site. So `siteCollection.archivalDetails.archiveStatus` (populated only for admin-archived TEAM sites, and only on `$select`) is the SECONDARY path; archived-OneDrive detection MUST cover the error branch (status 423 / code `resourceLocked` / an `~archived` message). Verified live 2026-06-01 against `lvmhfashion-my.sharepoint.com:/personal/m_lopez_us_celine_com`. Reactivation is admin-only (SharePoint admin center → Reports → OneDrive accounts → Reactivate) and billing-gated — unreachable on the delegated Teams token, so filtering (not "gaining access") is the only CLI-side fix.
Affects: `errorIndicatesArchived` in `src/domain/utilities/archive-status.ts`; any code that must distinguish an archived site from a merely missing one.

## [gotcha] 2026-06-01 | the machine-readable mutation report needs the `json` reporter added — `mutation.json` is otherwise stale

`stryker.conf.json` here emits `["clear-text", "progress", "html"]` — NO `json` reporter — so `reports/mutation/mutation.json` is NOT refreshed by a normal run; it retains whatever a prior run left behind, and parsing it (per the 2026-05-30 lesson) then reports STALE survivors as if current. To get an accurate machine-readable survivor set you must temporarily add `"json"` to `reporters` (alongside `"incremental": false`), run, parse `mutation.json` for `status === "Survived"`, then restore the config before committing. `reports/` is gitignored, so a stale `mutation.json` lingers silently with no hint it is out of date.
Affects: any scoped or full `bunx stryker run` where you parse `mutation.json` to close the 90% gate.

## [decision] 2026-06-01 | per-document sensitivity labels are unreachable on the Teams token (content scope absent)

Per-document MIP sensitivity labels read via the POST action `/drives/{d}/items/{i}/extractSensitivityLabels` (the `driveItem` resource exposes no GET label field in Graph v1.0). That action requires the delegated scope `InformationProtectionContent.Read.All`, which is NOT in the Teams web-client token's fixed grant — the token only carries `InformationProtectionPolicy.Read` (the *catalog* scope behind `list-sensitivity-labels`). So a per-item label read 403s like the other token-ceiling blocks (`Chat.Read`, `ChannelMessage.Read.All`, …). Spiked during the "full-fidelity document context" feature and DROPPED: the catalog command stays, but `download-drive-item-as-markdown --include-metadata` cannot surface a per-file `**Sensitivity:**` line on this token. The only fix is a custom Azure AD app (separate CLIENT_ID) granting the content scope — out of scope for the inherited Teams token.
Applies to: any future attempt to read/extract a driveItem's assigned sensitivity label; do not add such a command against the Teams token.

## [decision] 2026-05-30 | "maximum SharePoint a delegated user can see" = union of search-index + membership/activity vectors

There is no single delegated Graph API for "every site I can access" (Microsoft confirms it; `/sites/getAllSites` is application-only and not access-trimmed; `/me/followedSites` 403s on this token). The practical maximum is the UNION of two complementary commands, each catching what the other can't:
- `search-all-accessible-sites` — deep-pages `POST /search/query` `entityTypes:['site']` via `from`/`size` until `moreResultsAvailable` is false. This returns the FULL security-trimmed site index (incl. sites you can open but aren't a member of). Critically, `GET /sites?search=*` (behind `search-sharepoint-sites-by-name`) returns a single capped page with NO `@odata.nextLink` — measured 80 sites — whereas the Search API reported `total:142` and paged out 141. Always prefer `/search/query` paging over `/sites?search=` for completeness.
- `list-accessible-drives` — 7 vectors: `/me/drives`, joinedTeams→`/groups/{id}/drive`, memberOf(Unified)→drive, `/me/drive/sharedWithMe`, private/shared channels (`/teams/{id}/channels`→`filesFolder`), activity (`/me/drive/recent`+`/following`+`/me/insights/{trending,used,shared}`), and ALL libraries per site (path-addressed `/sites/{host}:/sites/{name}:/drives` — sites routinely have a non-default "Teams Wiki Data"/custom library the default-drive vectors miss).
Measured union on one tenant: 190 distinct site-roots vs 80 from `search *` alone (2.4×). Drive-level granularity matters — a site-root can host several driveIds; measure new coverage by driveId, not webUrl host/path (site-root dedup undercounts secondary libraries).
Applies to: any "show me everything I can reach in SharePoint/OneDrive" request on the Teams web-client delegated token.

## [gotcha] 2026-05-30 | Stryker `incremental` reports optimistically — `rm`-ing the file is not enough; set `incremental: false`

The 2026-05-29 note said `rm -f reports/stryker-incremental.json` fixes stale verdicts. It does not fully: with `testRunner: command` + `incremental: true`, a scoped `--mutate` run still over-credits kills (measured 85.4% where a truly-clean run was 84.0%) even after deleting the file. The CLI flags do not help — `--incremental false` is parsed as a config-file path ("Invalid config file 'false'") and `--no-incremental` is rejected as an unknown option. The only reliable accurate gate is editing `stryker.conf.json` to `"incremental": false` (optionally `"cleanTempDir": false` to keep the sandbox for inspection), running, then restoring. Restore both before committing — they are dev-speed defaults, not part of a feature change.
Affects: any scoped `bunx stryker run --mutate <file>` where you need a trustworthy pass/fail against the 90% break threshold.

## [gotcha] 2026-05-30 | a "survived" mutant you can't explain is usually an operand-level equivalent — pull the exact id

When a mutant survives that you're *certain* your tests kill, do not assume it's the whole-expression mutant. Stryker emits a separate mutant per operand of a compound condition. Example: `e.type === 'api_error' && e.status === 404` produces, among others, id=83 (whole→`true`, killed) AND id=86 (left operand `e.type === 'api_error'`→`true`, survived). id=86 makes the predicate `e.status === 404` — behaviorally identical because only the `api_error` member of the `GraphError` union carries a `status` field, so no representable input distinguishes it: a genuine equivalent mutant, correctly left alive. To diagnose: parse `reports/mutation/mutation.json` (`Bun.file().json()` — the HTML report embeds multi-file code-frames that break naive brace-slicing) for `{id, location.start.line, mutatorName, replacement, status}`, then activate the precise id in the sandbox via `__STRYKER_ACTIVE_MUTANT__=<id> bun test` and read the exit code. Don't write a test for an operand mutant the type system makes unreachable.
Affects: closing the last few points to a 90% mutation gate on files with compound boolean guards.

## [decision] 2026-05-29 | OOXML side-channel extraction composes on a shared `ooxml-*` core

docx/xlsx/pptx are all ZIP packages with byte-identical `docProps/{core,app,custom}.xml` and the same `*.rels` relationship graph. The metadata + image features share four format-agnostic modules: `src/infra/ooxml-zip-adapter.ts` (`openOoxmlZip`), `src/use-cases/commands/ooxml-xml-walker.ts` (fast-xml-parser traversal), `ooxml-metadata.ts` (core/app/custom props + all-`*.rels` external-link scan + VBA-macro flag), and `ooxml-metadata-to-markdown.ts` (render primitives). Per-format modules (`docx-metadata`, `xlsx-metadata`, `pptx-metadata`, …) own only body-specific extractors and compose the shared core. New OOXML work (ODF is the next zip-based candidate) must reuse these, not re-parse the zip.
Applies to: any future Office-document content/metadata feature.

## [decision] 2026-05-29 | `--include-metadata` appends for docx/xlsx, stands alone for pptx

The flag surfaces side-channel content (authored-but-not-rendered: properties, comments, tracked changes, speaker notes, hidden slides, defined names, VBA-macro presence). For formats with a convertible markdown body (docx via mammoth, xlsx via sheetjs) the metadata is appended as a `## … metadata` section. pptx has no markdown body in this CLI (it 415s toward `*-as-pdf`), so with the flag it returns the metadata as a *standalone* document and without it keeps the 415. Macro-enabled (`.docm`/`.xlsm`/`.pptm`) and template (`.dotx`/…) variants alias onto their base format via `office-extensions.ts`. New format metadata follows the same rule: append if there is a body, standalone if not.
Applies to: `office-to-markdown.ts`, `convert-mail-attachment-to-markdown.ts`, the `extract-*-images` family.

## [gotcha] 2026-05-29 | binary parts need a separate adapter from the metadata zip adapter

`ooxml-zip-adapter` pre-decodes every entry as a UTF-8 string (correct for XML, corrupting for binary images). Image extraction therefore uses a *separate* `src/infra/ooxml-media-extractor.ts` that returns raw `Uint8Array` per `*/media/*` part. Do not reuse the metadata adapter for binary content. The `--output-dir` flag mirrors `--output-path` (`persistMediaIfRequested` in `output-path.ts`) — single-file `--output-path` cannot split a `media` array across files.
Affects: any feature that pulls binary parts out of an OOXML package.

## [gotcha] 2026-05-29 | mutation: assert `error.type`, never guard on it

`if (!result.ok && result.error.type === 'api_error') expect(...)` lets a mutant that changes the error `type` skip the inner assertions — the test still passes, the mutant survives. This was the dominant cause of surviving mutants across the image feature. Assert the discriminant directly: `expect(result.ok).toBe(false); if (result.ok) return; expect(result.error.type).toBe('api_error');`. Also add a non-`io_failed` write-error case to kill the `type === 'io_failed' ? message : type` ternary survivors in `output-path.ts`.
Applies to: every `Result`-returning use-case test.

## [gotcha] 2026-05-29 | Stryker incremental cache lies after test-only changes; 1-element `.toSorted` is an uncovered fn

Two traps hit while mutation-testing: (1) `incremental: true` in `stryker.conf.json` serves *stale cached verdicts* when only test files change (finishes in seconds, scores don't move) and the report also lists cached unrelated files — `rm -f reports/stryker-incremental.json` (or scope `--mutate` to specific files) for an honest re-measure. (2) A `.toSorted((a,b) => …)` over a *single-element* array never invokes the comparator, so the comparator shows as an uncovered function even at 100% line coverage — drop the sort until there are ≥2 elements.
Affects: `bun run mutate:changed`, any manifest-derived list (`mediaProducingCommands` in `cli.ts`).

## [gotcha] 2026-05-06 | elevated capture deadline must accommodate federated-IdP interactive auth

The first elevated-token capture in `src/infra/browser-auth.ts` capped the poll deadline at 60 s, assuming silent SSO would always succeed via persistent profile cookies. On Okta-fronted tenants (e.g. ExampleCorp) the Okta passive WS-Fed session expires every few hours, after which `m365.cloud.microsoft/*` redirects through `okta.<tenant>.com/.../sso/wsfed/passive` and stays there pending interactive sign-in. 60 s is not enough — bump the deadline to the full `pollDeadlineMs` (5 min default) so the user has time to finish the Okta dance inside the visible Edge popup.
Affects: `src/infra/browser-auth.ts` `acquireElevatedToken`; any tenant where the Microsoft-Graph IdP is federated to a third party with shorter session lifetimes than Microsoft's own SSO.

## [decision] 2026-05-06 | use `m365.cloud.microsoft/search` as the elevated-capture URL

`https://m365.cloud.microsoft` (root) sometimes lands on `/chat/blocked` or stalls on bootstrap; `/search` is a more reliable target — it forces a Graph-aud token issuance for OfficeHome (`4765445b-...`) and M365ChatClient (`c0ab8ce9-...`) within the first 5–8 seconds of `domcontentloaded`. Both appids are on `ELEVATED_APP_IDS`, so either matches the filter. Empirically verified with a Playwright bearer-trace probe.
Applies to: `src/infra/browser-auth.ts` `M365_CLOUD_URL`.

## [decision] 2026-05-06 | historical-version content needs an ODSP-allow-listed Graph token

The 3 historical-version commands (`download-drive-item-version-content`, `download-drive-item-version-as-markdown`, `download-drive-item-version-as-pdf`) cannot use the Teams web client token alone — the embedded tempauth in the returned `streamContent` URL fails ODSP's `logicalPermissions` allow-list and 403s. Login captures a *second* Graph token from `https://m365.cloud.microsoft` whose first-party identity is M365ChatClient (`c0ab8ce9-e9a0-42e7-b064-33d422df41f1`), an app on the allow-list. New commands hitting ODSP-gated endpoints must call `graph.getBinaryElevated`, not `graph.getBinary`.
Applies to: any future command that 403s with `logicalPermissionAccessDenied`.

## [gotcha] 2026-05-06 | m365.cloud.microsoft refuses headless Edge

Initial elevated-token capture used Playwright in headless mode against `https://m365.cloud.microsoft`. Microsoft's anti-automation served an interstitial / token-issuance loop never fired. Switching `launchContext` to `headless: false` (visible Edge window) made the silent-SSO + Graph-fetch path work end-to-end. The persistent profile cookies still do the SSO; only the chrome needs to be visible.
Affects: `src/infra/browser-auth.ts` `acquireElevatedToken` flow; the brief Edge popup at token expiry is intentional.

## [gotcha] 2026-05-06 | bun text coverage reporter omits line numbers above ~99%

When `bun test --coverage` reports a file at 99.x% but the per-tier 100% gate fails, the text reporter sometimes prints an empty `Uncovered Line #s` cell, leaving you blind. Switch the bunfig coverage reporter to `lcov`, re-run, and grep `coverage/lcov.info` for `DA:N,0` lines to find the exact missing line numbers. Used to spot line 186 of `auth.ts` (a closing `}` of a nested `if`) — the fix was flattening the nested if so the closing brace stopped existing as a separately-instrumented line.
Rule for next time: don't add throwaway tests for impossible branches; reshape the source so the uncovered line stops being a distinct instrumented unit.

## [gotcha] 2026-06-10 | Bun bundler vs runtime CJS-default interop: dist/cli.js can crash where every test passes

A dep-importing infra adapter is NOT verified until `dist/cli.js` (the bundle) has executed the code path on a real input — under both `bun` and `node`. Bun's RUNTIME `import('@kenjiuno/msgreader').default` yields the class, so all 3600+ source-run tests passed; Bun's BUNDLER (node-mode `__toESM`) sets `.default` to the whole CJS exports object, leaving the class at `.default.default` — ".msg → Object is not a constructor" shipped broken in 1.4.0 for ~2 days. Packages with `exports.default = class` + `__esModule` are affected; `module.exports = class` (word-extractor) is immune. Fix pattern: `const Ctor = typeof d === 'function' ? d : d.default` with both shapes unit-tested. The QA playbook's A7 bundle-exec smoke now guards the class.
Affects: every lazy `await import()` of a CJS dependency in `src/infra/**`; `bun run build` consumers.

## [decision] 2026-06-10 | Non-Graph registry commands use the optional `executeLocal(fs, params)` field on Command

`convert-local-file` established the pattern: a registry command whose input is the local filesystem exports BOTH the registry-typed `execute` (returns a redirect error pointing library consumers at the local variant) AND `executeLocal(fs, params)`; `cli.ts` routes at the single execute call site (`cmd.executeLocal !== undefined ? executeLocal(fs, …) : execute(graph, …)`), reusing the whole option/validation/output-path pipeline. Register such commands in graph-scopes.test's COMMANDS_WITHOUT_SCOPES. Side gotcha: adding option `aliases` switches commander from `requiredOption` to `option` (required-ness moves to the zod schema), so commander-level "missing required option" tests must target an alias-free command.
Affects: any future offline/local command; `src/composition/cli.ts` execute dispatch.
