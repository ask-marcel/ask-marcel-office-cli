# QA Audit Playbook — full-surface health check

A repeatable, full-coverage audit of the `ask-marcel` CLI. Every run walks the **entire** command surface (all commands, all parameters, all global flags), both **offline** (gates, bundle, manifest consistency) and **live** (read-only probes against the real Microsoft 365 tenant). The output is a severity-ranked findings report — fixes happen only after the maintainer approves a plan ("report, then plan").

Run it with the `/qa-audit` skill, or hand this document to a fresh session and execute phase by phase.

---

## 0. Ground rules

**Prerequisites**
- Clean working tree on `main`, freshly built (`bun run build`) and locally deployed (`npm i -g .` — never `ask-marcel update`, which pulls npm `@latest` and clobbers local builds).
- Logged in (`ask-marcel login`); token cache valid at `~/.ask-marcel/token-cache.json` (includes the elevated M365ChatClient token for the 3 historical-version commands).
- `docs/commands.json` regenerated this run (`bun run docs:gen`) — every matrix below is driven from it.

**Safety rules (non-negotiable)**
1. **Read-only guarantee**: only GET/search commands against the tenant. The two POSTs (`microsoft-search-query`, `search-all-accessible-sites`) are searches — allowed. Nothing else mutates, but stay alert for any new command that does.
2. **No tenant content in committed files.** Raw findings (file names, subjects, IDs) go to the gitignored report dir (`.claude/qa-reports/`). Anything committed must be sanitized.
3. **Never fetch external URLs** (`@microsoft.graph.downloadUrl`, SharePoint CDN links) directly — the CLI follows redirects internally; agent harnesses often block those domains.
4. Quote every `b!…` drive-id in **single quotes** (zsh history expansion on `!`).
5. Use `--output-path` for any multi-MB byte fetch — never let base64 megabytes transit stdout/context.

**Severity rubric**
| Level | Meaning | Examples from past runs |
|---|---|---|
| **P1** | Active hazard: data leak, crash, silent corruption, context/cost bomb | 13 MB base64 leaking to stdout despite `--output-path`; `.msg` crashing in the bundled `dist/cli.js` |
| **P2** | Missing capability with a clumsy multi-step workaround | no mail-side zip handler (manual `unzip -O GBK` dance) |
| **P3** | Papercut: discovery cost, inconsistent naming/flags, unclear error | `--id` rejected where docs say "by ID"; opaque "too many arguments" on a flag typo |
| **R** | Removal candidate (see §G rubric) | a command fully covered by a consolidated sibling |
| **I** | Improvement (works, but worse than it should be) | error message that doesn't name the next command to try |
| **F** | Future / roadmap item | `pageCount` on PDF attachment metadata |

**Outputs of a run**
1. `.claude/qa-reports/YYYY-MM-DD.md` (gitignored) — full findings with repro commands and raw evidence.
2. A row appended to the **Run log** (§I, committed) — sanitized counts only.
3. A proposed fix plan (separate message, awaiting approval) for everything P1–P3.

---

## A. Baseline gates (offline)

| # | Check | How | Pass criteria |
|---|---|---|---|
| A1 | Tests | `bun test` | 0 fail |
| A2 | Lint, honest | `rm -f .eslintcache && bun run lint` | 0 errors AND 0 warnings (the cache can mask prettier warnings — always delete it first) |
| A3 | Typecheck (dev) | `bun run typecheck` | clean |
| A4 | Coverage | `bun run coverage`; check the **real exit code**, not a piped grep's | exit 0, "all files meet their tier gate". If a file shows >99% with no line numbers, switch bunfig to lcov and grep `DA:N,0` |
| A5 | Mutation | `rm -f reports/stryker-incremental.json && ./node_modules/.bin/stryker run stryker.conf.json` (node shebang — NOT `bun --bun x`, which breaks @babel/generator; delete the incremental file or it serves stale verdicts) | ≥90 overall; investigate any file <90 |
| A6 | Build (both tsconfigs) | `bun run build` | bundle + declaration emit clean. The dev typecheck does NOT run the build's declaration emit — dual-tsconfig CJS-interop errors only appear here |
| A7 | **Bundle execution smoke** | Run `dist/cli.js` under **both** `bun` and `node` on every vendored fixture: `convert-local-file --path` × {`src/test-helpers/assets/sample.msg`, `src/test-helpers/assets/legacy-sample.doc`, a temp `.csv`, a temp `.zip`, a missing path}, AND `extract-local-file-images --path src/test-helpers/assets/image-sample.docx` (the `producesMedia` family uses JSZip+unpdf — a separate bundler-interop surface; the image-bearing `.docx` is now **vendored** at `src/test-helpers/assets/image-sample.docx` → expect `ok`, `media=1`, `image/png`). The unpdf PDF-image branch is exercised live in D2 (mail-attachment re-probe). | Correct output under both runtimes. **`bun test` runs from source and can NEVER catch bundler-interop breakage** — the `.msg` "Object is not a constructor" bug shipped green through every gate and only this check finds its class |
| A8 | Deploy parity | `cmp "$(npm root -g)/ask-marcel-office-cli/dist/cli.js" dist/cli.js`; `ask-marcel --version` vs `package.json` | byte-identical, versions match |
| A9 | Dependency hygiene | `grep -E '"(latest|\*)"' package.json` (must be empty); `bun outdated` (note majors, don't bump) | no `latest`/`*`; outdated list recorded in report |

## B. Surface-consistency audit (offline, manifest-driven)

Drive everything from `docs/commands.json` so the checks scale to any future command count.

| # | Check | How | Pass criteria |
|---|---|---|---|
| B1 | Generated-docs freshness | `bun run docs:gen` then `git diff --exit-code docs/COMMANDS.md docs/commands.json` | no diff — the generated surface matches the registry |
| B2 | **Prose-docs truth sweep** (README.md + docs/USAGE.md + the lifecycle `--help` texts in cli.ts) — run-1 lesson: counting checks alone missed README omitting two shipped headline features and USAGE carrying four stale claims. Four sub-checks: **(a) numbers re-measured** — command counts, test count (`bun test` total), `--help` byte size, `help-json --terse` sizes, any "~N KB"/"N+ tests" claim; **(b) feature coverage** — diff the features shipped since the previous run-log row (commands added, capabilities like .msg/GBK/offline) against README's feature bullets: everything headline-worthy must appear; **(c) contract claims verified by execution** — stdout/stderr discipline, `--output-path` semantics, token-cache path+perms, exit codes: run the command, don't trust the prose; **(d) library-API section** — examples still compile against `src/index.ts` exports, special cases (e.g. `executeLocal`) documented | every claim re-verified or fixed in-run (doc fixes are exempt from report-then-plan — they ARE the audit's output) |
| B3 | **Phantom command references** | Extract every backticked `` `command-name` `` / `ask-marcel <name>` from all summaries, option descriptions, responseShapes, error-hint strings (grep `err({` messages in `src/use-cases/commands/`), **and from README.md, docs/USAGE.md, and the addHelpText blocks in cli.ts**, and verify each against the registry name list — flags too (run-1 found `resolve-mail-link --link` taught downstream; the real flag is `--url`) | zero references to non-existent commands or flags (meta.test validates per-command flags only — prose and downstream docs ship typos silently) |
| B4 | Example validity | For each manifest entry, run `ask-marcel <name> --help`; spot-parse each `example` string for flags that exist | all helps render; example flags all registered |
| B5 | Naming consistency | Group commands by family (`*-drive-item-*`, `convert-*`, `list-*`…); flag concept→name map (e.g. every message-scoped command uses `--message-id`; sole-id commands carry the `--id` alias) | inconsistencies listed (e.g. `download-onedrive-file-content` is the known odd one out — rename+alias is a standing F item) |
| B6 | Error actionability sweep | `grep -rn "err({" src/use-cases/commands/ --include="*.ts" \| grep -v test` — read every message | each error names the failing thing AND the next command/flag to try; flag the ones that don't as I items |
| B7 | help-json budget | `ask-marcel help-json \| wc -c` and `--terse --category <each> \| wc -c` | full manifest growth tracked run-over-run; terse per-category stays ~≤15 KB |
| B8 | Overlap / removal scan | For each pair of commands in the same category, ask: is A's output a strict subset of B's with the same inputs? List candidates under §G rubric | candidates documented with evidence, never removed in-run |

## C. Parameter matrix (every command × every option; live where needed)

Generate the worklist from the manifest (do this fresh each run):

```bash
bun -e '
const m = JSON.parse(await Bun.file("docs/commands.json").text());
for (const c of m.commands) {
  const opts = (c.options ?? []).map(o => `--${o.name}${o.required ? "*" : ""}${(o.aliases??[]).map(a=>` (alias --${a.name})`).join("")}`).join(" ");
  console.log(`${c.name} :: ${opts || "(no options)"}`);
}' > /tmp/qa-param-worklist.txt
wc -l /tmp/qa-param-worklist.txt
```

**Probe discipline (run-1 lesson):** every probe must parse the JSON envelope and assert `ok === true` explicitly — never `2>/dev/null`, never count array lengths without checking `ok` first. A missing-required-flag error envelope silently read as "0 results" nearly became a false P1 in run 1. Derive flags from the manifest, not memory — ~12 hand-typed flags were wrong in run 1 (all caught by the CLI's own error UX, which is itself evidence worth recording).

For **every command**, exercise each row of this table (live against the tenant for Graph commands — use IDs harvested in Phase E1; offline for `convert-local-file`, `resolve-*`, `docs`, `help-json`):

| Case | Expectation |
|---|---|
| All required flags, valid values | `ok:true` envelope matching `responseShape` |
| Each required flag omitted (one at a time) | error **names the exact missing flag** (commander `requiredOption`, or zod message for alias-carrying commands) |
| Each flag with empty string `""` | clear validation error, never a silent no-op (the `--output-path ""` class) |
| Each flag repeated twice | rejected by the no-repeat parser, not silent last-wins |
| Each alias spelling | behaves identically to canonical; validation errors reference the alias the user typed |
| Invalid value per typed flag (non-numeric `--top`, malformed `--filter`, bogus enum) | actionable validation/Graph error, no stack trace |
| Unknown flag (e.g. the `item--id` typo) | error + hint; note opacity as I item (standing: commander's "too many arguments" for bare-word typos) |
| OData flags where declared (`--select/--filter/--top/--orderby/--expand`) | honored or documented-as-ignored (mailboxSettings class); `--top` >1000 rejected client-side |
| Pagination (every `pagination: true` command) | `nextLink` hoisted to envelope top level; `next-page --url` round-trips |

**Global-flag cross-product** (run against representative commands of each shape, plus EVERY `producesBytes`/`producesMedia` command for the leak check). _As of 2026-06-15 (F-01 fix) `docs/commands.json` AND `help-json` both carry `producesBytes`/`producesMedia`/`mutates`, so derive the leak-scan worklist straight from the manifest: `jq '[.commands[]|select(.producesBytes or .producesMedia)|.name]' docs/commands.json`._:

| Flag | Check | Pass criteria |
|---|---|---|
| `--output-path` | **Stdout-leak scan**: every `producesBytes` command with `--output-path /tmp/qa/<n>` → measure stdout bytes | stdout ≤ 2 KB regardless of payload size; file lands byte-correct; envelope has `savedTo`, **no `base64`, no `contentBytes`** (the 13 MB P1 class) |
| `--output-path` | on a plain-JSON command | clear "did not return inlined bytes" error listing supported commands |
| `--output-path` | passthrough-vs-`.pdf` extension mismatch | `passthrough_extension_mismatch` guard fires |
| `--output-dir` | every `producesMedia` command | files land with flattened collision-free names; `base64` → `savedTo` |
| `--output json` / default text | error AND success paths under both | JSON envelopes parse; text mode renders |
| `--id` aliases | every command with a SINGLE required id-option (2026-06-16: broadened from mail-only to all sole-id commands) | works; commands with two required ids carry no `--id` (would be ambiguous) — guarded by a meta-invariant |

## D. Conversion deep-dive (fixtures + live)

The conversion pipeline is the CLI's core value. Two passes:

**D1 — Contract matrix (offline, fixtures).** Convert each format through every applicable entry point and compare against this expected-behavior table (update the table only on intentional behavior change — it IS the contract):

| Input | as-markdown (drive) | mail-attachment | drive-zip / mail-zip entry | convert-local-file | as-pdf |
|---|---|---|---|---|---|
| docx/docm/dotx | markdown (mammoth); `--inline-images`, `--include-metadata` | same | same | same | PDF |
| xlsx/xlsm | tables per sheet; `--max-cells` cap hint | same | same | same | PDF |
| pptx | `## Slide N` text | same | same | same | PDF (preferred for layout) |
| odt/ods/odp | markdown via content.xml | same | same | same | PDF |
| csv | markdown table (quote-aware, embedded newlines, `\|` escaped); cap hint over `--max-cells` | same | same | same | raw passthrough note |
| pdf (born-digital) | text layer → text/plain | same | same | same | raw bytes (no pdf→pdf) |
| pdf (scanned/no text) | 415 hint → vision model | same | note | same | raw bytes |
| .xls / .doc legacy | sheetjs table / word-extractor text | same | same | same | PDF (Graph renders) |
| .ppt legacy | 415 hint → convert-to-pdf first | same | note | hint incl. "upload to OneDrive" | PDF |
| .msg | H1 subject + From/To/Cc/Date + body + `## Attachments` recursed (depth-capped) | same | same | same | itemAttachment-style rejection |
| .zip | — | — | (nested archives noted, not unpacked) | files envelope; GBK/CP437 names decoded; >100 entries → `truncated` | — |
| png/jpg/gif/webp/bmp/tiff/ico | 415 image hint → vision | same | note | same | 415 image hint (mail) |
| svg | text/plain (XML sniffs as text) | same | same | same | — |
| valid-UTF-8 text, any/no extension | text/plain passthrough | same | same | same | raw bytes + passthrough tag |
| binary, unknown/no extension | 415 generic hint (`<no-extension>` case) | same | note | same | Graph attempt or clear error |
| loop/fluid/wbtx/whiteboard | Graph `?format=html` round-trip | n/a | n/a | **cannot** (documented) | n/a |

Cross-command consistency rule: the **same file** through drive / mail / zip / local must produce the same converter output. Sanctioned divergences (maintainer-approved 2026-06-10): (1) zip maps errors to notes; (2) unconvertible-input HINT WORDING is context-specific by design — top-level callers point at their own sibling commands, while files NESTED inside a container (zip entry, .msg attachment) always get the container-neutral `NESTED_HINTS` wording (QA-007). The converted CONTENT itself must never diverge.

**D2 — Real-tenant samples (live).** Via `search-onedrive-files` / `search-my-documents` / `list-mail-attachments`, locate ≥1 real instance of each format above (prior runs used real .msg, GBK-named vendor zips, scanned PDFs) and convert it. Synthetic fixtures miss real-world quirks: X500 sender addresses, compressedRtf-only bodies, TrueType hinting warnings, fonts with no `glyf` table.

## E. Live Graph drift probes (tenant)

Microsoft moves things. Each run, re-verify the full endpoint surface AND the known-fragile register.

**E1 — Category sweep.** One live happy-path call per command (full coverage — harvest IDs as you go: `list-drives` → drive-id → `list-folder-files` → item-id → …). Record HTTP status class per command in the report. Anything non-200 that worked last run = drift finding.

**E2 — Known-fragile register** (dated; extend as new fragilities surface):

| Probe | Why fragile | Last status |
|---|---|---|
| `list-teams-chat*` / `get-teams-chat-message` (chatsvcagg→csa substrate) | Microsoft-internal API, moved 2026-05 (`/api/csa/<region>/api/v{1,3}`) | healthy non-interactively 2026-06-15; **TIMED OUT 2026-06-16** (find-chats-with-user — chatsvcagg silent-SSO timeout). csa non-interactive health is **session/cookie-freshness dependent**, NOT stable — treat as warm-up-required, not drift; error guidance correct (logout&&login) |
| 3 `download-drive-item-version*` commands (ODSP allow-list, elevated token) | needs M365ChatClient-identity token from login's second capture | working 2026-06; BLOCKED 2026-06-15 AND 2026-06-16 (elevated-token silent-SSO timeout, QA-011) |
| `?format=pdf` input list (38 extensions) | CDN-side list drifts | working 2026-06-15 (docx→pdf live). **pdf→pdf returns PASSTHROUGH (raw bytes + `passthrough:true` + note), NOT a reject** — re-confirmed 2026-06-16 (convert-mail-attachment-to-pdf on a PDF) |
| `?format=html` input set (loop/fluid/wbtx/whiteboard only) | Office docs always `Sandbox_InputFormatNotSupported` | confirmed 2026-05; `.whiteboard` live-verified ok 2026-06-15 (no whiteboard harvested 2026-06-16 — carry) |
| OneNote reads on >5K-item doc libraries | Graph blocks ALL OneNote reads (tenant-side) | **personal notebooks UNBLOCKED 2026-06-15 and 2026-06-16** (all 7 cmds ok). 2026-06-16: the 5K-item limit fired on a SharePoint *site's* notebook library (list-sharepoint-site-onenote-notebooks) — CLI surfaced it with the correct actionable message (tenant-side, not CLI) |
| Elevated token capture in NON-interactive sessions (distinct from substrate) | silent SSO (IdP-federated) times out without a user at the browser → the *elevated m365.cloud.microsoft* path blocks list-chats/get-chat/list-chat-members/download-drive-item-version. The *substrate* csa token is USUALLY less affected but is ALSO session-dependent (timed out 2026-06-16 — see row 1). Error guidance correctly suggests `logout && login`. | observed 2026-06-15 AND 2026-06-16 (QA-011); plan audits with an interactive warm-up or accept the blocked subset |
| `/me/followedSites`, `/sites/delta` | always 403 on delegated; commands use `/sites?search=` | avoided by design |
| Archived sites | 423 `resourceLocked` on GET /sites/{id}; site-search probes & excludes | working 2026-06 |
| Elevated-token capture (visible Edge, IdP-federated SSO) | headless refused; federated IdP needs the 5-min deadline | working 2026-05 |

**E3 — Auth lifecycle.** `scopes-check` vs the scope map; expired-token behavior (clear re-login hint, no stack trace); `logout` → command → actionable "not logged in" error → `login` recovers. Verify no command exceeds the Teams-token scope ceiling (commands needing Chat.Read*, ChannelMessage.Read.All, Contacts.Read*, TeamMember.Read.All, TeamsTab.Read.All, Channel.Read.All cannot ship).

**E4 — Limits & failure modes.** One throttling observation if encountered (429 → message quality); one multi-MB attachment via `--output-path` (wall-clock + stdout size); folder-as-file, malformed id, foreign-tenant id → error quality.

## F. Agent-ergonomics review (the "improve" bucket)

Walk the canonical agent journeys end-to-end with fresh eyes, counting round-trips and wrong turns *caused by the CLI's docs/errors* (not by the agent). Journeys (extend over time):

1. "Read this Outlook thread and all its attachments" (message → list attachments → per-type conversion; zips, .msg, images).
2. "Find the X deck on SharePoint and summarize it" (site search → drive → search/browse → convert).
3. "What changed in this Excel file?" (versions → version content → diff).
4. "Read this local folder's report.docx and the vendor zip" (`convert-local-file`, no login).
5. "Who emailed me about Y last month?" (search-mail → read → resolve links in body).
6. A Teams-chat retrieval (substrate commands).

For each journey log: number of CLI calls, dead ends, which `--help`/`help-json` lookups were needed, every error encountered and whether its message self-healed the journey. Each friction point becomes an I finding. Also measure cold-start (`time ask-marcel --version`) — lazy-import regressions show here.

## G. Gap analysis & roadmap (the "missing / add / remove" bucket)

**Inputs to mine each run:**
- Plugin workaround docs (`../ask-marcel-plugin/**/CLAUDE.md`, `.claude/references/*`) — every workaround is a feature request (the P1/P2/P3 audit of 2026-06 came from exactly this). **Check the reverse direction too**: any plugin doc still teaching a workaround for behavior the CLI has since fixed is downstream staleness — flag it (and its phantom flags/commands) for a plugin-side doc update.
- This repo's memory/LESSONS deferred items.
- The Roadmap register below (carry-over).
- Graph API surface diff: for each category, list documented read-only endpoints with no corresponding command; judge usefulness, don't auto-add.

**Removal rubric (R findings)** — propose removal only when ALL hold:
1. Functionality fully reachable via another command (strict subset).
2. No plugin/doc references it.
3. Keeping it costs something real (manifest tokens, maintenance, user confusion).
Precedent: the 3 `download-drive-item-version-*` commands consolidated into one `--format` command.

**Roadmap register** (live list — append/strike each run):

| Item | Class | Origin | Status |
|---|---|---|---|
| `pageCount` on PDF attachment metadata (agents chunk reads at 20 pages) | F | plugin audit 2026-06 | open |
| Polymorphic `read-mail-attachment` (auto-route by content-type) | F | plugin audit 2026-06 | open |
| Rename `download-onedrive-file-content` → `download-drive-item-content` + alias | F/P3 | naming audit 2026-06 | open |
| meta.test: validate backticked command references against registry (phantom-name class — would have caught F-04) | F | QA 2026-06 | open (F-04 fixed manually 2026-06-15; general guard still wanted) |
| Friendlier commander error for bare-word flag typos (`item--id`) | I | session 2026-06 | open |
| Recycle-bin listing (metadata-only — no content API on delegated) | — | probe 2026-06 | dropped (Graph can't) |
| Cold-start reduction (lazier imports / precompiled registry) | I | QA-012 | **partial 2026-06-16**: lazy-loaded the two heaviest deps (xlsx 181 ms, mammoth 158 ms) via dynamic `import()` in their infra adapters → bundle init 603→530 ms, bun --version 0.65→0.51 s; conversions still self-contained (`--bin` intact). **Ceiling finding**: node cold-start (~0.9–1.0 s) is dominated by PARSING the 7.4 MB bundle, NOT init — dynamic import defers only init. The big win needs `--external` for the heavy deps (bundle 7.4 MB→1.2 MB) + lazy-loading winston too, which **breaks the compiled `build:bin` binaries** (they need everything bundled) → a distribution-level maintainer decision, not shipped here |
| Flaky time-boundary test helper: `chatsvcaggJwt()` recomputes `exp` from `Date.now()` and is evaluated twice (seed + expectation) so a sub-second delay desyncs `exp` → intermittent `bun test` failure (reddens A1 / pre-commit) | **P3** | QA-G16-01 (2026-06-16) | **FIXED 2026-06-16 (working tree)** — `chatsvcaggRequest(token?)` takes an explicit token; the 2 flaky tests capture one `chatsvcaggToken` for seed+assertion (matches the already-correct test at browser-auth.test.ts:1363). Four-check loop green (3775/0, lint/typecheck/coverage clean); test-only, no dist change |
| One shared `--drive-id` description constant | P3 | QA-003 | open (SharePoint pointer DONE 2026-06-10; 5 wording variants remain — minor) |
| `login` hang after silent token refresh | **P1** | QA-010 | **SHIPPED 2026-06-10 (cd6e74e)**; interactive race repro still opportunistic |
| chmod 0600 on token-cache.json at write | P3 | QA-001 | **DONE 2026-06-10 (72a6c70)** — verified 0600 live 2026-06-15 |
| Password-protected PDF honest message | P3 | QA-006 | **DONE 2026-06-10 (74fc486)** — verified live 2026-06-15 |
| Container-neutral hints for .msg/zip-nested attachments | I | QA-007 | **DONE 2026-06-10 (fe1cfd9)** |
| Orphaned `ask-marcel-temp` cleanup | P3 | QA-008 | **DONE 2026-06-10 (fe1cfd9)** |
| OData `--top` family alignment | P3 | QA-013 | CLOSED no-change (deliberate, documented) |
| SSO-timeout errors suggest `login` | I | QA-011 | CLOSED no-change — guidance present, re-confirmed live 2026-06-15 |
| Byte/media + `mutates` metadata serialized in commands.json AND help-json | P3 | QA 2026-06-15 (F-01) | **DONE 2026-06-15** |
| Read-only `--help` narrative derives writes-vs-searches from `mutates` (no mislabel) | P3 | QA 2026-06-15 (F-03) | **DONE 2026-06-15** |
| login help phantom `--use-playwright` removed | P3 | QA 2026-06-15 (F-04) | **DONE 2026-06-15** |
| `update-mail-draft --id` alias (family consistency) | P3 | QA 2026-06-15 (F-05) | **DONE 2026-06-15** |
| Local-only command errors stamp `source: cli` not `graph` | I | QA 2026-06-15 (F-02) | **DONE 2026-06-15** |
| Wrap raw JSZip error (no leaked docs URL) on non-zip input | P3 | QA 2026-06-15 (F-08) | **DONE 2026-06-15** |
| Stale doc numbers (README counts, help-json size hints) | I | QA 2026-06-15 (F-06) | **DONE 2026-06-15** |
| Vendor an image-bearing `.docx` fixture + add `extract-local-file-images` to A7 bundle smoke (producesMedia/JSZip+unpdf interop is un-smoked) | I | QA 2026-06-15 | **DONE 2026-06-16** — vendored `src/test-helpers/assets/image-sample.docx` (1.8 KB, real 1×1 PNG); A7 row now smokes `extract-local-file-images` on it under node+bun (ok, media=1, image/png). unpdf PDF-image branch live-confirmed in D2 (101 imgs); JSZip media path now in the offline bundle gate |
| Surface "use `--output-path` for large payloads" per byte command (get-mail-message-mime/get-mail-attachment default to inline base64) | I/F | QA 2026-06-15 | open (mitigated by F-01 producesBytes flag) |
| Pre-existing mutation debt — `docs.ts` (54%) / `docs-render.ts` (70%); `list-accessible-drives.ts` (89.58%); calendar `list-calendar-event-instances.ts` (88%) / `list-calendar-view-delta.ts` (86%) — all surfaced by full-file scoping | I | QA 2026-06-15 / -06-16 | open — **partial 2026-06-16**: added a real MAX_PAGES bound test (pagination can't loop forever). Remaining survivors are a STRUCTURAL CEILING: the `if (r === undefined) return` guards on `Promise.all` results are type-required (noUncheckedIndexedAccess) AND their false-direction mutant is equivalent (the guard never fires) — unkillable without deleting type-safe code. Overall scoped gate stays ≥90. Recommend accepting the ceiling, not adding contrived tests to game per-file scores |
| `--start`/`--end` aliases on calendar-view family (agents + plugin reach for them) | F | QA 2026-06-15 | **DONE 2026-06-16 (working tree)** — all 6 date-window commands (list-calendar-view/-delta/-event-instances, list-{specific,group,shared}-calendar-view) gain `--start`/`--end` → canonical key; live-verified (`--start/--end` → ok); meta-invariant + behavioral tests added |
| Downstream plugin doc sweep — stale `>/dev/null`/GBK-unzip workarounds + phantom flags (`--link`,`--start`,`get-event`,`--filter`,`--user-email`) across 13 files | — | QA 2026-06-15 | open (separate `ask-marcel-plugin` repo) |

## H. Report template

```markdown
# QA report — YYYY-MM-DD (vX.Y.Z, commit <sha>)
## Verdict: HEALTHY / DEGRADED / AT RISK
## Scores (per phase A–G: pass/warn/fail + one line)
## P1 / P2 / P3 findings
| ID | Sev | Area | Finding | Repro | Evidence |
## Regressions vs previous run (anything that passed last time and fails now)
## R candidates (with rubric evidence)  ## I improvements  ## F roadmap deltas
## Drift register updates (E2 table rows to change)
## Raw appendix (gitignored only): full command×status table, timings
```

Then — separately, after the report — propose the fix plan and **wait for approval** before changing any code.

## I. Run log (append a sanitized row per run; this table is the health history)

| Date | Version / commit | Auditor | P1 | P2 | P3 | R | I | F new | Verdict | Report |
|---|---|---|---|---|---|---|---|---|---|---|
| 2026-06-08 | 1.4.0 | plugin session (pre-playbook) | 1 | 1 | 1 | 0 | 2 | 2 | DEGRADED | plugin audit doc |
| 2026-06-10 | 1.4.0 / 444c4fc | Claude (first full playbook run) | 1 (carried: login-hang) | 0 | 5 | 0 | 6 | 6 | HEALTHY | .claude/qa-reports/2026-06-10.md — 173/173 cmds ALL live-verified (blocked subset closed after maintainer re-login); offline matrix 128×4 pass; leak scan clean; OneNote unblocked. Same-day fix wave shipped: QA-010 P1 + 001/002/003/005/006/007/008/009 fixed, 011/013 closed-as-correct (cd6e74e..fd7332b, scoped mutation 96.3%); full A5 aborted (pre-fix sandbox) — rerun next audit |
| 2026-06-15 | 1.4.0 / 53a4af3 | Claude (second full playbook run) | 0 (carried QA-010 shipped) | 0 | 6 | 0 | 2 | 4 | HEALTHY | .claude/qa-reports/2026-06-15.md — 176 cmds (was 173: +mail-draft writes, +image extraction, +Retry-After). Offline matrix 176×4 = 606 cases, 0 fail; byte+media leak guards clean; B6 error sweep clean. Live: 174 cmds, 158 ok / 6 boundary-err / 4 BLOCKED (QA-011 elevated-token) / 6 skip — **ZERO drift, ZERO P1 leaks** (4.78 MB→103 B stdout). Same-run fix wave (working tree, not yet committed): F-01 byte/media/mutates serialization, F-03 read-only narrative, F-04 login phantom flag, F-05 update-mail-draft --id, F-06 stale numbers, F-02 local err source, F-08 jszip-leak wrap. Gates: 3775/0 tests, lint/typecheck clean, coverage 100% all tiers, deploy byte-identical; added-line mutation 100%. A5 full run ~13 h (infeasible) — scoped-on-changed used. Improvements vs playbook: pdf→pdf now passthrough; Teams substrate healthy non-interactively |
| 2026-06-16 | 1.4.0 / 910b5c3 | Claude (third full playbook run) | 0 | 0 | 1 | 0 | 1 | 0 | HEALTHY | .claude/qa-reports/2026-06-16.md — 176 cmds, no product source change since 06-15 (only docs/style commits; all F-01…F-08 verified still in force). Offline matrix 606 cases 0 fail; offline byte+media+path guards clean. A5 scoped mutation **91.81% ≥90** (changed sort lines 100% killed; list-accessible-drives 89.58% = pre-existing debt). A7 bundle smoke OK both runtimes. Live: 176 swept → 144 ok / 12 boundary / 3 env-blocked (QA-011) / 11 NODATA / 5 skip / 1 tenant-5K — **ZERO CLI drift, ZERO P1 leaks** (156 KB→186 B; 101 imgs via unpdf leak-free). D2 PDF attachment: scanned→vision hint, pdf passthrough, **F-08 held live**. Personal OneNote unblocked; substrate timed out (session, not drift). **New P3: G16-01** flaky `chatsvcaggJwt()` time-boundary test helper (A1 3774/1, 5/5 green isolated — test-only). New I: list-accessible-drives mutation debt. Plugin staleness persists (13 files, separate repo). Post-approval fixes committed (26ac4fe..6f0e189): **G16-01** de-flake; **calendar --start/--end aliases** (6 cmds, live-verified, scoped mutation 93.70%); **vendored image-docx + A7 smoke**; **MAX_PAGES pagination bound test**. Deferred (maintainer decision): --id spread (contradicts message-id-only convention), cold-start refactor (QA-012), plugin-repo doc sweep |
