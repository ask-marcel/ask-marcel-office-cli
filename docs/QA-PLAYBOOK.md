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
| A7 | **Bundle execution smoke** | Run `dist/cli.js` under **both** `bun` and `node` on every vendored fixture: `convert-local-file --path` × {`src/test-helpers/assets/sample.msg`, `assets/legacy-sample.doc`, a temp `.csv`, a temp `.zip`, a missing path} | Correct output under both runtimes. **`bun test` runs from source and can NEVER catch bundler-interop breakage** — the `.msg` "Object is not a constructor" bug shipped green through every gate and only this check finds its class |
| A8 | Deploy parity | `cmp "$(npm root -g)/ask-marcel-office-cli/dist/cli.js" dist/cli.js`; `ask-marcel --version` vs `package.json` | byte-identical, versions match |
| A9 | Dependency hygiene | `grep -E '"(latest|\*)"' package.json` (must be empty); `bun outdated` (note majors, don't bump) | no `latest`/`*`; outdated list recorded in report |

## B. Surface-consistency audit (offline, manifest-driven)

Drive everything from `docs/commands.json` so the checks scale to any future command count.

| # | Check | How | Pass criteria |
|---|---|---|---|
| B1 | Docs freshness | `bun run docs:gen` then `git diff --exit-code docs/ README.md` | no diff — commands.json, COMMANDS.md, and the README's command counts all match the registry |
| B2 | README truth | Walk README claims vs reality: install steps, command counts ("N read-only", "N GET + 2 POST"), flag examples, feature claims | every claim verifiable |
| B3 | **Phantom command references** | Extract every backticked `` `command-name` `` / `ask-marcel <name>` from all summaries, option descriptions, responseShapes, error-hint strings (grep `err({` messages in `src/use-cases/commands/`), and verify each against the registry name list | zero references to non-existent commands (meta.test validates flags, NOT command names — a `search-sharepoint-site-files` typo ships silently) |
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

**Global-flag cross-product** (run against representative commands of each shape, plus EVERY `producesBytes`/`producesMedia` command for the leak check):

| Flag | Check | Pass criteria |
|---|---|---|
| `--output-path` | **Stdout-leak scan**: every `producesBytes` command with `--output-path /tmp/qa/<n>` → measure stdout bytes | stdout ≤ 2 KB regardless of payload size; file lands byte-correct; envelope has `savedTo`, **no `base64`, no `contentBytes`** (the 13 MB P1 class) |
| `--output-path` | on a plain-JSON command | clear "did not return inlined bytes" error listing supported commands |
| `--output-path` | passthrough-vs-`.pdf` extension mismatch | `passthrough_extension_mismatch` guard fires |
| `--output-dir` | every `producesMedia` command | files land with flattened collision-free names; `base64` → `savedTo` |
| `--output json` / default text | error AND success paths under both | JSON envelopes parse; text mode renders |
| `--id` aliases | every sole-message-id command | works; commands with two ids reject `--id` |

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
| `list-teams-chat*` / `get-teams-chat-message` (chatsvcagg→csa substrate) | Microsoft-internal API, moved 2026-05 (`/api/csa/<region>/api/v{1,3}`) | working 2026-06 |
| 3 `download-drive-item-version*` commands (ODSP allow-list, elevated token) | needs M365ChatClient-identity token from login's second capture | working 2026-06 |
| `?format=pdf` input list (38 extensions) | CDN-side list drifts; pdf→pdf rejected | working 2026-06-10 (docx→pdf live) |
| `?format=html` input set (loop/fluid/wbtx/whiteboard only) | Office docs always `Sandbox_InputFormatNotSupported` | confirmed 2026-05 |
| OneNote reads on >5K-item doc libraries | Graph blocks ALL OneNote reads (tenant-side) | **UNBLOCKED 2026-06-10** — notebooks/sections/pages/page-content all 200 live |
| Elevated + Teams-substrate token capture in NON-interactive sessions | silent SSO (IdP-federated) times out without a user at the browser → list-chats/get-chat/list-chat-members/list-teams-chat\*/find-chats-with-user/download-drive-item-version BLOCKED | observed 2026-06-10; plan audits with an interactive warm-up or accept the blocked subset |
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
- Plugin workaround docs (`../Ask-Marcel-Plugin/**/CLAUDE.md`, `.claude/references/*`) — every workaround is a feature request (the P1/P2/P3 audit of 2026-06 came from exactly this).
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
| meta.test: validate backticked command references against registry (phantom-name class) | F | QA 2026-06 | open |
| Friendlier commander error for bare-word flag typos (`item--id`) | I | session 2026-06 | open |
| Recycle-bin listing (metadata-only — no content API on delegated) | — | probe 2026-06 | dropped (Graph can't) |
| Fix `login` hang after silent token refresh (browser teardown; stderr "waiting on user" line) | **P1** | plugin doc 2026-06-08, confirmed open QA 2026-06-10 | open |
| chmod 0600 on token-cache.json at write (docs already claim it) | P3 | QA 2026-06-10 (QA-001) | open |
| One shared `--drive-id` description constant (11 commands lack the SharePoint pointer) | P3 | QA 2026-06-10 (QA-003) | open |
| Password-protected PDF: honest dedicated message | P3 | QA 2026-06-10 (QA-006) | open |
| Container-neutral hints for attachments nested in .msg/zip | I | QA 2026-06-10 (QA-007) | open |
| One-time cleanup + prevention for orphaned `ask-marcel-temp` (un-dotted) root folder | P3 | QA 2026-06-10 (QA-008) | open |
| Align OData `--top` support within command families (list-site-columns vs list-site-content-types; microsoft-search-query page-size) | P3 | QA 2026-06-10 (QA-013) | open |
| Cold-start reduction (~1.2 s node: lazier imports / precompiled registry) | I | QA 2026-06-10 (QA-012) | open |
| SSO-timeout errors should suggest `ask-marcel login` explicitly | I | QA 2026-06-10 (QA-011) | open |

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
| 2026-06-10 | 1.4.0 / 444c4fc | Claude (first full playbook run) | 1 (carried: login-hang) | 0 | 5 | 0 | 6 | 6 | HEALTHY | .claude/qa-reports/2026-06-10.md — 173/173 cmds, 146 ok + 4 by-design + 7 SSO-blocked + 14 nodata; offline matrix 128×4 pass; leak scan clean; OneNote unblocked; Stryker full run pending at close |
