# Workflow

The durable plan, the four-check loop, coverage gates, lint discipline, and the editor/CI rules that keep them enforced. Run through this after every code change; nothing ships until it is clean.

## The durable plan (`.claude/PLAN.md`)

Before a multi-step task, write the plan to `.claude/PLAN.md`, not just to the chat. Chat context is lost the moment the session ends or a fresh one starts; a committed file survives both. The plan is the resumability contract: a returning human or a cold Claude session reads it first and continues at the same place with the same information, instead of re-deriving the plan from a half-remembered thread.

**What it holds.** The goal in a sentence or two; the whole-task definition of done; the ordered steps, each with a checkbox and a per-step DoD (the concrete check that proves the step is finished); and a short breadcrumbs section (paths touched, commands to rerun, decisions made and why). Enough that a reader with zero prior context could pick up the next unchecked step.

```markdown
# PLAN: <task>
Status: in progress. Started YYYY-MM-DD.

## Goal
<one or two sentences>

## Definition of done (whole task)
- <checkable outcome> ...

## Steps
1. [x] <step>  DoD: <check> [met]
2. [ ] <step>  DoD: <check>
3. [ ] <step>  DoD: <check>

## Notes / breadcrumbs
- <decisions, paths, commands a cold reader needs>
```

**Lifecycle.**
- **Write it before executing** a multi-step task; trivial one-step work does not need it.
- **Keep it live.** Tick each box the moment its DoD is met; mark steps done / in-progress / blocked as you go. The on-disk file is the source of truth, current even between commits, so it survives a context loss immediately.
- **Commit it alongside the work slices** it describes (it rides with the same commits, not a separate noisy stream), so a fresh clone has the current plan.
- **Close it out at task end.** All boxes ticked, or a note on what remains for next time. When the next task begins, overwrite it.

**PLAN.md is not LESSONS.md.** `PLAN.md` is the *mutable current plan* and is rewritten and overwritten freely. `.claude/LESSONS.md` is *append-only memory* (decisions, gotchas), rewritten only by an approved compaction pass (`references/lessons.md`). A durable decision that outlives the task graduates from a PLAN breadcrumb into a `[decision]` lesson; the plan step itself is transient. See `references/lessons.md`.

**On resume.** Start of session, read `.claude/PLAN.md` (alongside the lesson files). If it shows an unfinished task, continue from the first unchecked step rather than re-planning. If the user's new request supersedes the open plan, say so in one sentence and overwrite it.

**Within a long run.** The live plan is also your context-budget checkpoint, not only a crash-recovery file: keeping it current means a long agentic run degrades gracefully instead of hitting a context limit blind, because the next step and its DoD are always on disk. When a task is too large to finish in one context window, decompose it into independently-checkpointed steps (or subagents) rather than driving one context past the wall.

## The four-check loop (after every change)

```bash
bun run test       # must pass: the randomized script (rule 36), never bare `bun test`
bun run lint       # 0 errors AND 0 warnings
bun run typecheck  # tsc --noEmit, clean
bun run coverage   # per-directory thresholds pass
```

If any of the four fail, fix the cause and re-run all four. Do not move on while one is red. Warnings have repeatedly hidden real issues (silent precedence bugs, dead returns, suppressed scanners); the zero-warning rule is not cosmetic.

The `package.json` scripts behind the four commands (`lint`, `lint:strict`, `typecheck`, `coverage`, plus the mutation and staged-lint ones) are printed once, in `references/bun-typescript.md` § `package.json`.

`bun run lint:strict` (~25 s) sets the env var `LINT_STRICT=1`; the same `eslint.config.js` reads `process.env['LINT_STRICT']` and conditionally adds a type-aware block (`parserOptions.projectService: true` plus `@typescript-eslint/no-unnecessary-type-assertion` and `@typescript-eslint/prefer-promise-reject-errors`). One config file, two modes, no separate `eslint.strict.config.js` to keep in sync. CI runs the strict version as a merge gate; the pre-commit hook runs the fast, non-type-aware `lint:staged` on the staged files, so run `bun run lint:strict` yourself in the inner loop to see the type-aware findings before you push.

## Zero warnings; no inline ignores

`bun run lint` is considered failing if it reports any warnings, not just errors. Two acceptable ways to clear a finding:

1. **Refactor the code** so the rule no longer fires. This is the default. If Snyk's string-literal-adjacent-to-key heuristic flags `const apiKey = 'sk-...'`, build the value at runtime from env vars. If `no-await-in-loop` fires, restructure the loop into `Promise.all` or accept the sequential cost with a targeted severity change.
2. **Configure rule severity at the project level** in `eslint.config.js`, with a comment explaining why. Reserved for rules that produce only false positives in this codebase's idioms: branded types, string-literal unions, bounded regexes, `security/detect-object-injection` on `Record<K, V>` lookups. For security-plugin rules, the two-part test below must be satisfied before disabling. Atelier already does this with the unicorn rules that contradict the standard or prettier (each named with its reason in `references/bun-typescript.md`), `security/detect-object-injection`, `security/detect-unsafe-regex`, and `security/detect-non-literal-fs-filename`.

**Forbidden everywhere, no exceptions:**

- `// eslint-disable`, `// eslint-disable-next-line`, `// eslint-disable-line`
- `// @ts-ignore`, `// @ts-expect-error`
- `// deepcode ignore`, `// snyk-ignore`, `// sonar-ignore`, `// istanbul ignore`
- `// prettier-ignore`, `// Stryker disable`, `// NOSONAR`, `/* c8 ignore */`, `/* v8 ignore */`, `// biome-ignore`, `// oxlint-disable`
- Any equivalent from another tool

If a rule needs suppression for a single line, the suppression is a lie: either the rule is wrong for this codebase (change severity at the project level) or the code is wrong for this codebase (refactor). A per-line suppression hides both.

The lint enforces this in both TypeScript configs (hard rule 15, since 2026-09-08). `linterOptions.noInlineConfig: true` makes every ESLint directive comment inert and reports it, so `--max-warnings=0` fails on the comment and the violation it hid surfaces beside it; `@typescript-eslint/ban-ts-comment` rejects every `@ts-` form, a described `@ts-expect-error` included; core `no-warning-comments` rejects the markers other tools read (the list above) anywhere in a comment. A tool that gains a new marker joins that list, and the smoke tests plant each form and require the red.

### Project-level rule disabling: a two-part test

Before adding a rule to the `'off'` list, both of these must be true:

1. **Every fire in this codebase is a false positive for our idioms.** Not "most", not "the current ones". Branded types + `Record<K, V>` lookups, bounded regexes with documented inputs, `chmodSync(mkdtempSync(...))` inside FS-adapter tests: these never represent a real exploit in the atelier style. If even one fire out of ten is a genuine finding, leave the rule on and refactor the other nine.
2. **The pattern the rule would catch if it fired correctly is something the production code cannot produce.** `security/detect-non-literal-fs-filename` matters for `node:fs` calls in production; atelier production code uses `Bun.file` instead (which the rule does not watch), so disabling globally loses nothing on the real attack surface. If production *could* produce the pattern, disabling masks a real vulnerability class.

Never disable a rule globally just to silence a single test, a single commit, or a single file. If the fire is localised, the right tool is a narrower `files:` scope in the ESLint config (disable the rule for `**/*.test.ts` only, for example): still at the project level, still with a comment, never inline.

## Keeping `coverage-preload.ts` in sync (auto-regeneration)

`scripts/coverage-preload.ts` lists every file in `src/{infra,composition,presenter}/` so the coverage table can include them at 0% if untested. Keeping this file in sync by hand is tedious and the failure mode is silent: a missing import means the new file never appears in the coverage report and the gate trivially passes.

One script handles this, shipped in the skill at `assets/`:

| Script | Job |
|---|---|
| `assets/regenerate-coverage-preload.ts` | Walks `src/{infra,composition,presenter}/`, excludes `*.test.ts` / `ports/` / `index.ts`, writes a fresh `scripts/coverage-preload.ts` with `/` in every path on every platform |

`scripts/coverage-preload.ts` itself is always **generated**, never hand-written: copy the regenerate script into a new repo and run it once to create the initial preload.

**Two modes:**

```bash
# Write a fresh preload from the current src/ tree.
bun run scripts/regenerate-coverage-preload.ts

# Exit non-zero if the on-disk file is out of sync, for pre-commit / CI.
bun run scripts/regenerate-coverage-preload.ts --check
```

Both modes fail, writing and comparing nothing, when the walk found files under a scan directory that the preload would not import. That was the Windows failure until 2026-10-09: `path.relative` answers `src\infra\x.ts` there, every path missed its group, the preload imported nothing, and `--check` still reported "in sync". Paths are now normalised to `/` where they enter. An empty scan directory is not a failure, since the bootstrap creates `src/presenter` and `src/composition` empty. `--selftest` proves the guard and the normalisation with Windows path semantics on any machine.

**Wire `--check` into CI, right before the coverage gate.** An out-of-sync preload silently lies about coverage, so the check belongs beside the gate it protects, and the coverage gate runs in CI:

```bash
# In .github/workflows/ci.yml, before the coverage step:
echo "pre-flight: coverage-preload sync" >&2
bun run scripts/regenerate-coverage-preload.ts --check
```

It is <100 ms and O(files), so it is cheap enough to also drop into the pre-commit hook as a pre-flight if you like, but its real home is CI, next to the coverage gate. Then the workflow becomes: add a new file under `src/infra/`, run `bun run scripts/regenerate-coverage-preload.ts`, stage both files together, commit. The `--check` invocation catches any forgotten regeneration before it merges.

## SDK-bridge lines: how coverage handles unreachable wiring

Some lines in `src/infra/**` exist solely to bridge to a third-party SDK and are structurally unreachable without launching the SDK for real. Concretely:

- `await import('playwright')` inside a closure that the smoke test never actually triggers
- `google.drive({ version: 'v3', auth })`, instantiates a real Google client; can't be exercised without real credentials
- `new MongoClient(url)`, opens a real driver
- A `process.on('SIGTERM', ...)` handler that the test runner never sends

These lines are covered by the **production-wiring smoke test** described in `references/testing-infra.md` whenever possible: the smoke test calls `createX(realDeps)` with a placeholder, which exercises the wiring line and asserts the resulting port has the right method shape.

When even the smoke test cannot reach a line (a closure inside a method that requires real SDK behaviour to enter), accept it as exempt. The 80% gate on `src/infra/**` is calibrated for this; a single adapter at 88-95% line coverage with the rest of the file fully tested is healthy.

**Rule of thumb.** Lines that exist solely to bridge into a third-party SDK (dynamic imports inside a closure, real-factory pass-throughs, SDK-instantiation one-liners) are exempt from line-coverage when the file's other paths bring it above the per-tier gate. Do not lower the gate; do not add a per-file skip in `bunfig.toml`. Just accept that the bridge line is the cost of doing business with the SDK and the gate is permissive enough to absorb it.

If a single file is dragged below the 80% gate by SDK-bridge lines alone, the right move is usually to refactor: split the bridge into a thinner `createX(realSdk)` that only does the instantiation, and a fatter `createXFromApi(api)` that holds all the logic. Then the bridge file is one or two lines (still uncovered, but tiny) and the logic file is fully tested.

## Coverage gates (per-tier, enforced by custom script)

Bun's built-in `coverageThreshold` is a single global number. It cannot express "100% on the domain, 80% on infra, skipped on test-helpers". The repo enforces per-tier rules via `scripts/check-coverage.ts`, which runs `bun test --coverage`, parses the text report, and applies path-prefix rules.

A ready-to-copy `check-coverage.ts` lives in the skill at `assets/check-coverage.ts`. It exposes `COVERAGE_RULES` and `SKIPPED` as top-of-file constants so tuning per-project takes a one-line edit.

| Path | Threshold (functions & lines) |
|:---|:---|
| `src/domain/**` | **100%** |
| `src/use-cases/**` (including `ports/`) | **100%** |
| `src/composition/**` (env.ts AND build-deps.ts) | **80%** |
| `src/presenter/**` | 80% |
| `src/infra/**` | **80% from day one**, not "once tests exist" |
| `src/test-helpers/**` | skip during normal runs (audit periodically, see below) |
| `src/main.ts` | skip (entry point; verified by integration) |

**`build-deps.ts` is no longer skipped.** The earlier policy excluded it as "composition root, verified live, no logic worth unit-testing". That was hedging. The composition root becomes fully unit-testable when (a) every "where do I read state from" point (file path, env var, system clock) is parameterisable, and (b) every "what do I write to / log to" sink can be injected as a port. See `references/architecture.md` (Composition root testability) for the optional-config-DI pattern.

Every `src/infra/*.ts`, `src/composition/env.ts`, and `src/presenter/cli.ts` carries a real 80% gate: most end up at 100% once the three infra-test patterns (see `references/testing-infra.md`) are in routine use. The "we'll add infra tests later" road leads to a coverage gate that trivially passes.

`bun run coverage` exits non-zero if any file falls below its gate and prints the offending paths with current-vs-required numbers. A tier summary at the end highlights the worst funcs/lines per tier, so a single sloppy file is visible without scrolling the per-file table.

It also exits non-zero on a file under `src/` that no tier and no skip rule claims: a new directory gets a tier in `COVERAGE_RULES` or, for a genuine non-code entry, a line in `SKIPPED`, never silence. It reads the table with `/` on every platform, because Bun prints `src\domain\x.ts` on Windows, where until 2026-10-09 every row matched no tier and the gate passed with "no files" in each. `bun run scripts/check-coverage.ts --selftest` feeds planted rows (a Windows row under its gate, a `src/` row in no tier) through the same parser and checks, and fails unless each is caught.

If a file cannot hit the gate, the fix is usually **restructure the code so the dead branch goes away**, not lower the threshold. A threshold reduction must be justified in the commit message.

### The coverage preload (mandatory)

`bun test --coverage` only reports rows for files the test runner imports. Untested infra files (no `*.test.ts`, no test imports them) are silently absent from the table, which makes the per-file gate trivially pass. Every adapter you forgot to test becomes invisible instead of failing loudly.

The fix is a preload file that side-effect-imports every infra, composition, and presenter module, so they appear in the coverage table at 0% if no test exercises them. Generate it with `assets/regenerate-coverage-preload.ts` (§ Keeping `coverage-preload.ts` in sync, above), never hand-write it. The generated file looks like:

```ts
// scripts/coverage-preload.ts
// Auto-generated by scripts/regenerate-coverage-preload.ts. Do not hand-edit.
// This file forces every module that belongs under a coverage gate to appear
// in `bun test --coverage` output, even when no test imports it. Without this,
// an untested adapter is silently absent and the per-file gate passes.

import '../src/infra/logger.ts';
import '../src/infra/sheets-google.ts';
import '../src/infra/telegram-http.ts';
// ... one line per infra / composition / presenter file
import '../src/composition/env.ts';
import '../src/presenter/cli.ts';
```

**Wire it at coverage time only, NOT in `bunfig.toml`.** `scripts/check-coverage.ts` spawns:

```bash
bun test --coverage --preload ./scripts/coverage-preload.ts
```

Do not put `preload = [...]` under `[test]` in `bunfig.toml`. The preload pulls in heavy runtime deps (e.g. `googleapis`, `winston`, `twitter-api-v2`, `@ai-sdk/google`) that would add 1–2 seconds to every plain `bun test`. Loading it only at coverage time keeps the inner-loop fast without losing the gate.

**Maintenance rule:** every new file in `src/infra/`, `src/composition/`, or `src/presenter/` must be added to `coverage-preload.ts` in the same commit: run `bun run scripts/regenerate-coverage-preload.ts` and stage both files together. Enforcement is mechanical, not goodwill: the shipped `ci.yml` runs `regenerate-coverage-preload.ts --check`, so a forgotten regeneration blocks the merge; wiring the same check into the hook as a pre-flight, to catch it at commit time, is optional (§ Keeping `coverage-preload.ts` in sync, above).

### `bunfig.toml`: minimal, no `coverageThreshold`, no `preload`

The global `coverageThreshold` in `bunfig.toml` must be **absent** when the per-tier script owns enforcement. If set, Bun exits non-zero on the global threshold before the script runs, the script's first check (a non-zero test status) bails with "`bun test --coverage` exited non-zero; fix test failures first", and the per-file violation breakdown never prints, so a threshold miss reads as a test failure.

`preload` must also be absent under `[test]`, see above.

Correct `bunfig.toml`:

```toml
[test]
coverage = true
coverageSkipTestFiles = true
coverageReporter = ["text"]
# NOTE: no `coverageThreshold` here, per-tier enforcement is in
# scripts/check-coverage.ts.
# NOTE: no `preload` here, the coverage preload is loaded only by
# `bun run coverage` (via --preload on the spawned bun-test command),
# so plain `bun test` runs stay fast.
```

When introducing a per-tier coverage gate in an existing repo, remove any global `coverageThreshold` *and* any `preload` from `bunfig.toml` in the same change.

## SonarLint findings caught at lint time

SonarLint runs IDE-side; CI and pre-commit do not see it. To keep IDE-only findings from drifting back in, ESLint is wired to catch them at lint time.

In `eslint.config.js` (one config, two modes, `LINT_STRICT=1` switches on the type-aware block). The canonical, complete flat config is printed once, in `references/bun-typescript.md` (§ `eslint.config.js`): `sonarjsPlugin.configs.recommended` is on, the type-aware lane adds `@typescript-eslint/no-unnecessary-type-assertion` and `prefer-promise-reject-errors`, and the six SonarJS rules turned off are each justified beside the switch there (the three dated ones, 2026-08-29 and 2026-09-05, are re-probed weekly by the skill repository's canary). Copy from that file, never from memory.

The conditional block runs only when `process.env['LINT_STRICT']` is set, so the inner-loop `bun run lint` skips it entirely. `bun run lint:strict` is just `LINT_STRICT=1 eslint`.

### Complexity gate (rule 35)

`complexity: ['error', 10]` in the base rules block (both variants) caps cyclomatic complexity at 10 per function: nine sequential guard clauses pass (complexity 10), ten fail (11), and so does a one-line chain of ten `&&` terms, which the size caps in SKILL.md (10 lines, one indentation level) never see. That is why the cap exists beside the size caps rather than instead of them, and why `sonarjs/cognitive-complexity` stays off: one metric, the one the size caps cannot substitute for. The fix is never a bigger number; split the function or replace the chain with a dispatch map (`references/complexity.md`). The Java variant enforces the same cap through PMD (`references/java-quarkus.md`, Gates and hooks). Both smoke tests plant a complexity-11 function and see the gate red.

### Common SonarJS findings and how to fix them

| Sonar ID | Symptom | Fix (never suppress) |
|:---|:---|:---|
| **S4325** | `x!` non-null assertion, or `x as Type` without real narrowing | Replace with a guard clause: `const found = xs.find(...); if (!found) throw new Error(...); return found;` |
| **S6594** | `"abc".match(re)` used for captured groups | Use `re.exec("abc")`: more efficient, avoids the global-flag trap |
| **S4123** | `await` on a matcher chain that is not a real `Thenable`, e.g. `await expect(p).rejects.toThrow()` | Use the `captureRejection(promise)` helper: see `references/result-type.md` |
| **S6551** | `String(err)` in a catch block | Use the shared `formatError(err: unknown): string` from `src/domain/utilities/format-error.ts` |
| **S6671** | `Promise.reject(value)` where value is not an `Error` | Change to `Promise.reject(new Error(...))`. For tests that deliberately reject with a non-Error, use a tiny `async (v: unknown) => { throw v }` helper |
| **sonarjs/void-use** | `void unusedParam;` to silence unused-var warnings | Drop the parameter entirely from the implementation. TypeScript's function-type **parameter contravariance** means a function with fewer parameters is assignable to a function-type with more. |

### Types must not lie

`Record<K, V>` says "every key maps to V". JavaScript's runtime says otherwise: missing keys return `undefined`. If the key set is open (user IDs, row IDs, environment variables), the honest type is `Partial<Record<K, V>>`.

```ts
// BAD - the type lies
type SheetRow = Readonly<Record<string, string>>;
const value = row['maybe-absent']; // typed as string, actually undefined

// GOOD - the type tells the truth
type SheetRow = Readonly<Partial<Record<string, string>>>;
const value = row['maybe-absent'] ?? ''; // typed as string | undefined, narrowed at use
```

Every consumer that did `value !== undefined` on the first form was calling a check the type said could never be false. The `Partial` form makes the check real.

## Trunk-based development

Integrate to one branch, `main` (the trunk), continuously. Commit straight to it, or through a branch that lives **less than a day** and merges back small. Long-lived feature branches are the thing this model exists to avoid: every day a branch diverges, the eventual merge gets riskier, review gets coarser, and `main` stops reflecting reality.

The rules already in this skill are precisely what makes committing to the trunk safe: trunk-based development is their reason for existing, not a separate concern:

- **Every commit keeps `main` releasable.** The fast pre-commit hook (below) blocks the obvious breakage locally (size, dependencies, secrets, identity, the discipline tripwires, staged lint, types) in a few seconds, and CI holds the full line (the whole test suite, coverage, mutation, strict lint) as the required merge check, so what reaches the shared trunk is green.
- **Commits stay small** (gate 1: ≤10 files AND ≤300 lines). Small commits are the unit of continuous integration; they review in minutes, revert cleanly, and bisect precisely. A 300-line ceiling is a trunk-based ceiling.
- **History stays linear and legible** (Conventional Commits, `commit-msg` hook). A trunk read top-to-bottom is the changelog.
- **Incomplete work hides behind a flag, not a branch.** When a feature spans several commits, keep each commit green and the half-built path dark behind a feature flag or simply unreferenced, never park weeks of work on a divergent branch. This is the same instinct as YAGNI and "minimal": ship the smallest safe increment.

Practical loop: pull/rebase often to stay close to the trunk; run the four-check loop after every change; when green, propose the commit and commit it once the user confirms (SKILL.md hard rule 25, the agent never commits or pushes on its own initiative); then push on their say-so. If a change is too big to land safely in one ≤300-line commit, split it into a sequence of green commits, not a long-lived branch. Releases are cut from the trunk (tag or release branch at the moment of release), never developed on for weeks beforehand.

This is the default for this codebase. It overrides any tooling habit of "branch first by default": branch only when a short-lived branch genuinely helps (e.g. a PR-review gate your team requires), and merge it the same day.

### Branch lifecycle (hard rule 38)

A branch is a short detour from the trunk: its whole life fits in a day, and it leaves nothing behind. `main` is the trunk and the host's default branch; a new repository starts on it (`git init -b main`). Six steps, the gates and the host holding steps 2, 3, 4 and 6:

1. Start from a freshly fetched `main` (`git fetch origin && git switch -c <name> origin/main`), or commit to `main` directly.
2. Keep your own branch current by rebasing it onto `main` (`git rebase origin/main`), never by merging `main` into it.
3. Land it by rebase or fast-forward: the host's "Rebase and merge", never a merge commit, and never a squash of several commits, which builds one commit that fails gate 1 on `main`. `check-commit-range.sh` rejects a merge commit in every pull request and every push (GitHub's synthetic merge of a pull request into its base, which a pull_request checkout sits on, is stepped over).
4. Delete it the moment it lands, on the remote and locally (`git push origin --delete <name>`, `git branch -D <name>`). For the agent the remote delete is a push, so it waits for the user's yes (rule 25); when the host refuses it, say so and point the user at the host's delete button.
5. Follow-up work after a landing starts a fresh branch from the new `main`, never more commits on the landed one.
6. Nobody force-pushes or deletes `main`: the host's ruleset refuses both, to admins too.

The host enforces what git alone cannot, and its owner sets it once, as code. `assets/check-host.py --apply`, run with an admin token after the first push, makes `main` the default branch (renaming the current default when there is no `main`, and printing the four commands each clone runs after the rename), allows rebase merges only, deletes head branches on merge, and creates the ruleset in `assets/host-ruleset.json`: on the default branch, with no bypass, no deletion, no force push and linear history, so no merge commit reaches `main` by any path. The ruleset requires neither a pull request nor a status check, because step 1 allows committing to `main` directly; a team that lands only through pull requests adds those rules to the file, and the check then holds them too. `--apply` changes the host's configuration, which is the owner's call: the agent proposes it and runs it on the owner's yes, like a push (rule 25).

```bash
python3 scripts/check-host.py --apply   # the owner, once, with gh signed in as an admin
```

`assets/branches.yml` runs two watchdogs every morning. `check-branches.sh` fails on a remote branch whose work is already on `main`, judged by content with `git merge-tree` so a rebase or squash merge counts as landed, and on one whose oldest commit not on `main` is more than a day old (`MAX_BRANCH_AGE_HOURS`, default 24). `KEEP_BRANCHES` (default `^release/`) exempts the release branches cut from the trunk when a release ships; a team that keeps other long-lived branches names them there, and owns the drift that follows. `check-host.py` fails when the default branch is not `main`, when `main` lacks a rule the ruleset file lists, or when a merge setting has drifted. It needs read access only: GitHub's REST API shows a read-only token none of the four merge settings, GraphQL shows all four (probed 2026-10-15), and a setting GitHub does not show is an error, never a pass. Neither is a merge gate: a branch goes stale while nobody pushes to it, and a setting drifts while nobody looks.

Two cases bend the host check. A private repository on GitHub Free cannot hold a ruleset, since rulesets need a public repository or a paid plan: its owner sets `HOST_RULESET: unavailable` in `branches.yml`, and the check names `main`'s rules as unchecked without reading them (GitHub answers that read with a 403 on Free too) and holds the default branch and the merge settings alone; on a repository that is not private the opt-out itself fails. What the plan withholds, nothing on the host enforces: a force push or deletion of `main`, a merge while CI is red (required status checks need branch protection), a push made with `--no-verify`, which CI reports only after it lands. Pro or Team, then `check-host.py --apply` again, gives the protection back. And a history rewrite that must force-push `main` (rule 26's scrub, Commit identity below) sets the ruleset's enforcement to disabled for that push and back to active after it; the morning check fails in between.

### Landing on green

A pull request lands itself. Each variant's CI workflow ships a `land` job that runs after every other job of the workflow passed (`needs:`; `check-workflow-assets.sh` fails a `land` job that misses one, or that runs on `always()` or `failure()`), and calls `assets/land-on-green.sh`. Nobody presses merge, so nothing lands red, which also closes the gap a private repository on GitHub Free leaves open, where the host cannot require a check. The script:

1. reads the merge commit CI checked out (`GITHUB_SHA`): its first parent is the base CI tested, its second the head;
2. waits up to `LAND_WAIT_MIN` minutes (default 20) for every other check on the head, a dependency audit in its own workflow or a review bot included, since `needs:` sees only the jobs of its own workflow, then refuses unless each passed, was skipped or is neutral;
3. refuses when `main` moved since CI ran. A merge made with `GITHUB_TOKEN` starts no push run on `main`, so landing onto a newer base would put a tree on `main` that no run tested. Rebase onto the new `main` and push; CI runs again and lands it;
4. rebase-merges the pull request pinned to the tested head (`gh pr merge --rebase --delete-branch --match-head-commit`), which deletes the branch on the remote as step 4 above asks.

A refusal is a red `land` check that names its reason. Three things skip the job: a draft pull request (marking it ready runs CI again, since the workflows listen for `ready_for_review`), one from a fork (its author has no write access, and its code would run with a write token), and the repository variable `LAND_ON_GREEN` set to `off`, for a team that wants a human review before each merge. Adoption (atelier-review-me's adopt mode) keeps the job out of the copied workflow until its gates block, since a gate held back would let it land code that fails the standard. The job runs the pull request's own copy of the script with a write token, which every same-repository author can already push; a fork never reaches it.

Rule 25 is unchanged: an agent still asks before it commits or pushes. Pushing a branch and opening its pull request is now the last human decision before `main`.

## Confirmation gates (rules 24 and 25)

Two behavioural gates, not lint-enforced; the discipline is the enforcement, exactly as for rule 11.

**Rule 24: never touch a test without explicit user confirmation.** Test files (`*.test.ts` by project convention, `*.spec.ts` too if a repo uses it) are confirmation-gated. Do not create, edit, rename, move, delete, skip (`.skip`, `.only`, `xfail`), or weaken (loosen an assertion, change an expected value, comment out a case) any test without first showing the user the exact test or diff and getting an explicit yes. Tests are the contract and the safety net; silently editing a failing test to make it pass, or deleting an inconvenient one, is the most dangerous move an agent makes, because it disables the very check that catches regressions. This holds under TDD (rule 11): the loop stays test-first, and the Red step becomes propose the failing test, get confirmation, then write it. When a test fails, the default is to fix the production code; changing the test is a last resort that needs the user's sign-off and a one-line reason. If asked to "just make the tests pass", never weaken them silently: surface the conflict and ask. The same applies to a change in `src/test-helpers/**` that would alter what existing tests assert.

**Rule 25: never commit or push on your own initiative.** Producing and staging the change is the agent's job; deciding to commit it is the user's. Even when the tree is green, even when a commit is the obvious next step, even mid-flow: stop, show what would be committed (the staged-diff summary and a proposed Conventional Commits message), and wait for an explicit yes before running `git commit`, and the same for `git push`. Do not infer "commit" from a general "do it" or "go ahead" on the task; the commit needs its own confirmation. An explicit "commit and push X" is that confirmation; silence is not. Rule 23 governs the message format, rule 24 the tests; this rule governs when a commit happens: only on the user's say-so.

**Unattended (headless) runs.** Both gates hold when nobody can answer. The one carve-out is creation: writing a NEW test for new code proceeds without the pause, because blocking on a question nobody can answer would make TDD impossible. Every other gated action on an EXISTING test (edit, weaken, delete, skip, rename) stays forbidden unattended, and so do commit and push; do the work, stage it, and put the gated proposals in the final report.

## Reply gate (the Interaction section)

SKILL.md's Interaction section says how a reply to a person reads, and five of its rules are mechanical: no em dash, no cut word, no bold lead-in on a list item, no decorative emoji, sentence-case headings. A doctrine line alone does not hold them: on 2026-10-04, 831 replies from 16 real cloud sessions in seven consumer repos carried a bold lead-in in 63 percent of the current model's long replies, and in 52 percent even where a tool call named atelier.

**The gate.** `assets/check-reply.py --hook` is a Claude Code `Stop` hook, wired by `assets/claude-settings.json`: copy the script to `scripts/check-reply.py` and the settings to `.claude/settings.json`, or merge their `Stop` entry when that file already exists. When the agent ends a turn, Claude Code sends the hook the reply (`last_assistant_message`; the transcript's last reply on a version without that field). A doctrine finding exits 2 with each tag, an example and its fix on stderr; Claude Code hands that text back to the agent, which sends the reply again, fixed. The second stop carries `stop_hook_active` and passes, so a reply is restated at most once and the gate cannot loop. Input that is not a Stop event exits 1, a visible hook error and never a silent pass. It needs `python3` and nothing else. It reads the event and writes its feedback as UTF-8 on every platform: Windows Python decodes a pipe as cp1252 unless UTF-8 mode is on, and until 2026-10-09 an em dash or an emoji arrived there as other characters and passed.

**What it costs.** A blocked reply appears twice, the original and its restatement, with a "Stop hook error" notice between them, and the restatement costs its own length in output tokens (a three-item list and its restatement: $0.037 together, 2026-10-04). The probe counts its firing from session logs, per arm: the gate's blocks, the ones after a session's first, the restatements that still broke a rule, and the replies that broke one mid-turn (below). The first reading (2026-10-11) covers one session, this repository's own: 4 blocks in 25 judged replies, all bold lead-ins, 3 of them after the session's first, every restatement clean. It holds against an explicit request too: asked for bold labels, the agent restated without them, the same rewrite-to-comply the hard rules ask of code.

**What it never sees.** The gate judges a turn's last message only. Text written mid-turn, such as a summary before an AskUserQuestion call, reaches the person unjudged: in the session of the first reading, 2 of the 6 replies that broke a rule after the gate went live got through that way. Holding the question does not close that gap. In a probe (2026-10-11), a PreToolUse hook that blocked the call with the same feedback saw the model retry the call with no text, twice, and fix the prose only in its final reply, after the call had run. So the gap is measured before anything is built for it: the probe's gate line counts the mid-turn replies that broke a rule (3 in that session, each a summary before a question, one of them from before the gate went live).

**What stays a count.** The file is also a probe: given files or directories (eval transcripts, `~/.claude/projects/<slug>/` session logs) it prints every finding and counts per arm. Its Simplified Technical English candidates (a hedged result such as "should pass", any hedge, an event passive, sentences over 25 words, paragraphs over 6 sentences, a question buried in a report) never block: the 831-reply reading found no habit to fix, and a rule allowing only can, must and will would push honest uncertainty toward false certainty.

**The proof.** Each variant's smoke test copies the pair as a bootstrap does and proves exit 2 with the `bold-lead-in` tag on a planted reply, a pass on a plain one, the loop guard, and that the copied settings run the copied script. In the skill repository, CI runs the file's `--selftest` (every tag on its own plant, both modes, the cut list against this section) and `check-workflow-assets.sh`, which fails a variant whose checklist does not copy the hook's script.

## Commit identity (rule 26)

Every commit carries an author and a committer (each a name plus an email), taken from git config, and whatever they are becomes permanent public history the moment you push. Carrying the contributor's real identity in that metadata is normal, the default of the whole open-source world, and never a finding, an audit item, or a publish blocker.

File contents are the opposite. No tracked file ever names a person, an employer, or a client: not a name in a comment or a LICENSE holder line, not an employer's internal hostname in a config, not a client name in a fixture. A content mention outlives the commit that added it, travels with every copy and quote of the file, and once pushed cannot be removed by anything short of a history rewrite. Where a holder or author string is structurally required, use a neutral handle (e.g. `atelier`). Host control files whose format is identities (CODEOWNERS, `.mailmap`) are metadata in file form, not mentions; they are exempt. The cheap moment to catch a mention is review (atelier-review-me checks it); the expensive moment is after a push.

**The gate.** `assets/check-identity.sh` is a tripwire in the shape of the discipline guards below, and a core gate: every shipped hook runs it on the staged added lines and every shipped CI workflow on the whole tracked tree (`--all`). It looks for the committer's multi-word git name in both orders and the email, under `--all` for every author and committer in the history, and for every entry of `IDENTITY_DENYLIST` (employer and client names; an environment variable, never a tracked file, since a tracked denylist would itself name what the rule forbids). A one-word git name is a handle and is skipped, as is a GitHub noreply address; CODEOWNERS and `.mailmap` are exempt; a lone first name stays a review duty. The fix for a hit is a neutral handle, never a suppression.

**Public blocks, private warns.** A name in a public repository's files is public the moment it is pushed; in a private one it is a debt, paid before the repository ever goes public. So the gate blocks a hit in a public repository and only warns in a private one, and the warning still lists each hit and names the rule (under GitHub Actions also as a `::warning::` annotation). A private repository declares itself: `.github/visibility`, committed, one line reading `private`. The hook reads that file; CI passes the host's own answer as `REPO_VISIBILITY` (from `github.event.repository.private`, in the identity step of every shipped CI workflow), and that answer wins: a file saying private on a public host fails CI with or without a hit, so a repository made public cannot keep warning. No file, or any other word in it, counts as public. Before a private repository goes public, delete the file, make `check-identity.sh --all` pass, and scrub from history the names the warnings let in (below).

Secrets are the other real pre-publish concern: run `gitleaks git` (the history-wide mode, not the pre-commit `git --staged`) before the first push to a public host. Secrets in history are always findings; metadata identities never are.

**Scrubbing pushed history is a rewrite, gated and user-initiated.** A one-time, destructive operation; never run it unprompted (rule 25). Use `git filter-repo` (install: `brew install git-filter-repo`): `--replace-text` removes a mention from file contents across history, and `--mailmap` remaps commit metadata when the user wants that changed too:

```bash
# replacements.txt, one rule per line (mention ==> neutral replacement):
#   Old Name==>atelier
#   old@company.com==>atelier@users.noreply.github.com
git filter-repo --replace-text replacements.txt --force
git remote add origin <url>          # filter-repo strips the remote as a safety measure
git push --force-with-lease origin main
```

To also remap the author and committer fields, add `--mailmap` with `Intended Name <intended@email> <old@email>` lines. `filter-repo` rewrites every commit SHA, so this is a coordinated force-push: anyone holding a clone must re-clone. Rule 38's ruleset refuses a force push to `main` from anyone: set its enforcement to disabled for this push and back to active after it (Branch lifecycle above).

**A force-push does not purge the old commits.** The rewritten branch no longer points at them, but the host keeps unreferenced commits reachable by their SHA, through cached views, and via any fork or open PR, until it garbage-collects on its own schedule. Treat a leaked commit as exposed even after the fix: rotate anything that was a live secret, and for a hard guarantee delete-and-recreate the repo or ask the host's support to purge.

## Gates: a fast pre-commit hook plus the full set in CI

The gate set has two homes, split by speed and not by importance (canon 15.1). The pre-commit hook runs the **fast gates** only, because a multi-minute hook trains `git commit --no-verify` (canon 15.3). Every gate, fast and slow, also runs in **CI**, the line that cannot be skipped and the required merge check (canon 4.6). The full test suite, per-tier coverage, and Stryker mutation are slow and grow with the codebase, so they live in CI and only in CI.

This is the **Bun-script variant's** mechanism. The Next.js monorepo uses `simple-git-hooks` (pre-commit runs gate 2, the identity gate and the discipline wrapper, then each package's test + lint; commit-msg runs commitlint) instead, see `references/nextjs-monorepo.md`. Never install both: `core.hooksPath` and `simple-git-hooks` overwrite each other. The Java variant's hook is `assets/pre-commit-java` (`references/java-quarkus.md`, Gates and hooks).

**The pre-commit hook (fast gates, target under ~5s), `assets/pre-commit`:**

| # | Gate | Purpose | Typical time |
|:--:|:---|:---|:--:|
| 1 | `scripts/check-commit-size.sh` | <=10 files AND <=300 lines | <1s |
| 2 | `scripts/check-package-json.sh` | no `"latest"` / `"*"` / bare dist-tag (rule 19); no foreign lockfile, no `scripts` entry calling `node`, `npm`, `npx`, `pnpm`, `yarn` or `vite` directly (rule 5) | <1s |
| 3 | `gitleaks git --staged --pre-commit` | secret scan on the staged diff, inline allow comments ignored | ~50ms |
| 4 | `scripts/check-identity.sh` | no person, employer or client named in the staged lines (rule 26; a warning in a private repository) | <1s |
| 5 | `scripts/check-disciplines.sh` | the core discipline tripwires on the staged lines: personal data channels, IO deadlines, data lifecycle (rules 27, 29, 30; Discipline tripwires below) | <1s |
| 6 | `bun run lint:staged` | ESLint on the staged TS files only | ~1-2s |
| 7 | `bun run typecheck` | `tsc --noEmit` clean | seconds |

Every gate here is O(staged files) or O(1). Typecheck is the one that grows with the whole codebase; if it exceeds the hook budget on your repo, move it to CI too. Never add the test suite, coverage, or mutation to the hook.

**CI (`assets/ci.yml`, the authoritative merge gate):** `check-commit-messages.sh` and `check-commit-range.sh` straight after checkout (both need only git history), then install on a frozen lockfile, `check-package-json.sh`, `gitleaks git` (full history, inline allows ignored; the workflow installs its own pinned, checksum-verified gitleaks), `check-identity.sh --all` and `check-disciplines.sh --all` (the whole tree), `lint:strict` (type-aware, zero warnings, ~25s), `typecheck`, `bun test --randomize` (the whole suite, rule 36), `regenerate-coverage-preload.ts --check`, `bun run coverage` (per-tier), `check-docs.sh` (the README's Verify block, `references/governance.md`), `bun run mutate:changed` on every event (the changed files only, 1-3 min per file; on a push the range is `github.event.before..HEAD`, which the workflow exports). The full sweep is never a commit gate: `assets/mutation.yml` runs `bun run mutate` once a day on a schedule. Make it a required status check in branch protection (canon 13.2) so a bypassed hook is still caught. The CVE scan is deliberately NOT in this job; it ships as its own workflow, `assets/audit.yml` (see Dependency CVE scanning (CI) below).

### Install once per clone

The copy steps live in one place, the Bootstrap checklist of `references/bun-typescript.md` (step 14): the hooks, every script they and the three workflows call, the workflows themselves, and `git config core.hooksPath .githooks`. The Bun smoke test replays that checklist, so it is the copy block that is proven to work; do not keep a second one here. Then generate the initial preload once: `bun run scripts/regenerate-coverage-preload.ts`.

`core.hooksPath .githooks` picks up **both** `.githooks/pre-commit` (the fast gates, on the staged diff) and `.githooks/commit-msg` (Conventional Commits, on the message), one config, two hooks. `.github/workflows/ci.yml` is the authoritative gate set that runs every gate on every push and pull request.

Install gitleaks (optional but recommended): `brew install gitleaks` on macOS, or grab a binary from `github.com/gitleaks/gitleaks/releases`. The hook degrades gracefully if `gitleaks` is missing (it warns and continues) so first-time clones don't break.

The `git config core.hooksPath .githooks` is the one step that is easy to forget. Without it, Git looks in `.git/hooks/` and your commit goes through unchecked. Document it in the repo's `README.md` install section.

### Commit size limits (gate 1)

`scripts/check-commit-size.sh` enforces the rule "≤10 files **AND** ≤300 lines (insertions + deletions)", i.e. it blocks any commit that exceeds *either* threshold. The limits are conservative because they force the discipline; loosening them undermines the rule.

Why:
- Small commits are easier to review, revert, and bisect.
- Large commits hide bugs (one slip across 300 lines is hard to spot).
- Every commit on `main` becomes git history that the next engineer reads: keep each one a coherent slice.

When working on a feature, **commit as you go**: one focused slice at a time. The gate is the safety net, not the policy.

### Dependency hygiene (gate 2)

`scripts/check-package-json.sh` blocks any commit where `package.json` declares a version as `"latest"` or `"*"`, and, since 2026-09-08, any commit that tracks or stages a `package-lock.json`, `npm-shrinkwrap.json`, `yarn.lock` or `pnpm-lock.yaml`, or whose `scripts` call `node`, `npm`, `npx`, `pnpm`, `yarn` or `vite` directly (rule 5; an env prefix and a segment after `&&` count, `bunx vite` does not). Every entry under `dependencies`, `devDependencies`, and `peerDependencies` must use a concrete version (`X.Y.Z`) or a real range (`^X.Y.Z`, `~X.Y.Z`, `>=X.Y.Z`).

Why:
- `"latest"` and `"*"` are non-deterministic. `bun install` on different days gives different `node_modules/` trees. The lockfile only partially mitigates this.
- The literal string `"latest"` semantically signals "always upgrade": a silent-break footgun that can pull in a major version change between two checkouts of the same commit.
- You don't audit what you didn't expect to install. Hidden upgrades from `"latest"` are how supply-chain attacks land.

Workflow:

- **Adding a package.** `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev). Bun resolves the actual latest version at install time and writes it as `^X.Y.Z`. **Never hand-edit `package.json` to add a dep**: the gate may pass on a manually-typed `^1.2.3`, but you lose the auto-pinning convention and the muscle memory drifts.
- **Bumping every dep to current latest.** Run `bun update`. This rewrites the existing `^X.Y.Z` ranges to the latest matching versions and updates `bun.lock`. Commit both files in the same change. Do this on a deliberate cadence (start of a release, dependabot-style cron, etc.), not silently on every commit.
- **Bumping one specific dep.** `bun update <pkg>` for a constrained bump, or `bun add <pkg>@latest` to force the absolute current latest into the same `^X.Y.Z` slot. Either way, no `"latest"` ends up in the file.
- **Initial scaffold.** When using the skill's `package.json` skeleton (in `references/bun-typescript.md`), the version ranges are samples. Run `bun install` to resolve them, then `bun update` to pull each dep to its current latest, then commit both files together. Verify with `bash scripts/check-package-json.sh`.

The gate reads the version strings inside the four dependency blocks only (`dependencies`, `devDependencies`, `peerDependencies`, `optionalDependencies`, one-line or multi-line), catching `"*"`, `"latest"`, bare dist-tags (`beta`, `alpha`, `next`, `canary`, `rc`), and an `npm:` alias resolving to one; real ranges like `^1.2.3` and `>=4.0.0` pass, and a dist-tag elsewhere in the manifest (`publishConfig.tag`) is not a finding.

### Secret scanning with gitleaks (gate 3)

The hook runs `gitleaks git --staged --pre-commit --ignore-gitleaks-allow --redact --verbose --no-banner`. Two distinct gitleaks modes, pick the right one:

- **`gitleaks git --staged --pre-commit`**: scans the staged-but-not-committed diff. Fast (~50 ms). Blocks re-introduction of secrets *before* they enter history. Use in pre-commit hooks.
- **`gitleaks git`**: scans the entire git history (every commit, every file ever). Slow. Use for periodic audits or CI checks. **Does not** belong in a pre-commit hook.

`protect` and `detect` are the pre-8.19 names of the same two modes, hidden and deprecated since. Every call carries `--ignore-gitleaks-allow`: without it an inline `// gitleaks:allow` comment on the line silences the finding, an inline suppression by another name (rule 15), and the ESLint configs reject the comment itself (`no-warning-comments`). A real false positive goes in a committed `.gitleaksignore` by fingerprint, with a reason in the commit.

Run `gitleaks git` once before the first push to GitHub to catch anything that snuck in pre-hook.

### Mutation testing with Stryker (a CI gate)

[Stryker](https://stryker-mutator.io/) generates small "mutants" of the production code (e.g. `>` becomes `>=`, `&&` becomes `||`, `return x` becomes `return undefined`) and runs the test suite against each. A mutant that survives means your tests don't actually pin the behaviour they appear to.

The atelier policy: **every file under `src/domain/**` or `src/use-cases/**` must score ≥90% mutation score** before it merges. CI is the enforcing home (`mutate:changed` on every pull request and push, the changed files only; the full `mutate` sweep runs once a day from `assets/mutation.yml`, never on a commit); locally, `bun run mutate:changed` (the files changed since your last push) and `bun run mutate:staged` are optional pre-checks. The threshold is the `break` value in `stryker.conf.json`.

The config ships as `assets/stryker.conf.json`, copied verbatim by the Bootstrap checklist (`references/bun-typescript.md`); it is the one copy, read it there. The choices that matter: the command runner runs `bun test --randomize --bail src/domain src/use-cases` (rule 36; the next paragraph says why the flags), `mutate` is `src/domain/**` and `src/use-cases/**` less tests and `ports/`, `thresholds.break` is 90, and `ignorePatterns` keeps non-source directories out of the sandbox with its file globs anchored to the root (`/*.json`, not `*.json`: gitignore semantics make a bare glob match at every depth, which drops a test's JSON fixture and fails the dry run).

There is no first-party `@stryker-mutator` Bun runner today (community plugins exist, but we don't depend on them), so we use the command runner: Stryker shells out to the test command once per mutant, and that command's cost is the whole run's. It ran the full suite to the end (~7 s a mutant on a typical codebase) until 2026-10-23. Two flags cut it without changing a verdict. `--bail` stops at the first failing test, and one failure is all a kill needs. The `src/domain src/use-cases` filters run only the core's tests, because mutants live only there; a mutant only an adapter test could kill now survives, which is the stricter reading of the 90 (the core's own tests pin its behaviour). On a scratch core of 34 mutants beside a 3 s adapter suite, a run went from 26 s to 1 s with the same 34 kills, and a weakened test scored 58.82 under both. A core with no test fails the dry run rather than passing. There is no `concurrency` key: Stryker's default uses every core above 4 less one, else all of them, so a 4-core CI runner keeps its 4 and a laptop uses what it has. `incremental: true` caches per-mutant results so unchanged code is not re-tested. `packageManager: "npm"` is needed because Stryker probes for a JS-ecosystem package manager and does not yet recognise Bun's lockfile. `ignorePatterns` skips non-source dirs from the sandbox copy, `.claude/` in particular often contains a symlink Stryker cannot copy (ENOTSUP).

**Three commands, three scopes:**

- **`bun run mutate`**: full run on `src/domain/**` + `src/use-cases/**`. Slow (1-2 hr on ~150 files). The daily scheduled sweep (`assets/mutation.yml`, plus `workflow_dispatch` after a large refactor), never a commit gate.
- **`bun run mutate:changed`**: in CI, the files the pull request or push changed; locally, the files changed since your last push (the branch's `@{push}`, or `origin/main` on a branch never pushed). Both add uncommitted edits and **untracked files**. Locally it is optional: CI's run over the whole pull request is the gate, so the local run checks only the work CI has not seen yet (until 2026-10-23 it compared against `origin/main` and re-mutated every file the branch had touched, on every push). Override the base ref with `BASE=HEAD~3 bun run mutate:changed`, or `BASE=origin/main` for the whole branch.
- **`bun run mutate:staged`**: files staged for the next commit. An optional local pre-push check; CI is the enforcing home, running `mutate:changed` on every pull request and push. Skips with exit 0 when no relevant files are staged, so commits to docs, tests, or scripts are unaffected.

**What `mutate:changed` guarantees, and the one thing it cannot.** Each guarantee exists because its absence produced a green run over unmeasured code:

- **Untracked files are in scope.** A brand-new `src/domain/order.ts` that was never `git add`-ed appears in no diff, so a scope built from diffs alone reports "nothing changed" and exits 0. The script unions in `git ls-files --others --exclude-standard`. (`mutate:staged` deliberately does not: a file must be staged to be committed, so the staged variant has no equivalent hole.)
- **An unresolvable base ref fails loudly.** `BASE` that does not exist makes every diff fail, and the `|| true` guarding the scope pipeline turns that into "no files changed" plus exit 0. The script now resolves `BASE` up front and exits 1. In CI this matters more than locally: `assets/ci.yml` checks out with `fetch-depth: 0`, but a consumer repo on a shallow checkout has no `origin/main` and would otherwise get a vacuously passing mutation gate on every pull request.
- **The base ref is refreshed and printed.** `origin/main` is a *local cache* of the remote, moved only by a fetch. Against a stale one, `$BASE...HEAD` still contains long-pushed commits, so the scope fills with files nobody touched. The script fetches when `BASE` is a remote-tracking ref (`|| true`, so offline degrades gracefully; `MUTATE_NO_FETCH=1` opts out) and prints the resolved base as short SHA, relative date, and ahead count, which is what makes staleness visible when you override `BASE` or opt out.
- **Every file holds the threshold, not the run's total.** Stryker's `thresholds.break` judges the whole run's score, and under `--mutate` the incremental report keeps earlier results for every other file in that score: a file at 55.56 passed in a run that totalled 90.00 (found 2026-10-23). So the run ends with `scripts/check-mutation-score.ts`, which reads the JSON report the run just wrote and holds each file the run mutated to the same `break`, read from the report; a file it was handed but the report lacks fails by name. `bun run mutate` ends with it too, over every file in the report.
- **Verdicts are fresh.** The run passes `--force`. Stryker's incremental cache keys on source-file hashes, so a test-only change (a stronger assertion, same source) replays a stale score. Prefer `--force` to deleting the incremental file: `incrementalFile` is owned by `stryker.conf.json`, so a config change would silently turn a hardcoded `rm -f reports/stryker-incremental.json` into a no-op.
- **The CI range is the event's range.** Under Actions the base is the pull request's base branch or, on a push, `github.event.before` (the zero SHA of a new branch and an unresolvable SHA fall back to `HEAD~1`, and a repository's first push, with no `HEAD~1`, mutates every file in scope), the same resolution as the commit gates. A push to main has HEAD == origin/main, so the old default would have tested nothing and exited 0. `BASE=` still wins.
- **Greenfield, before the first push:** there is no `origin/main` yet, so pass a local base (`BASE=$(git rev-list --max-parents=0 HEAD)`) or skip the command until the remote exists. Failing loudly is the point; do not paper over it by restoring a silent exit 0.

**The part no script can encode:** `mutate:changed` output is only a statement about push scope if the ref behind it is current. Fetch before reading any `origin/main`-relative output that way, and confirm what is actually unpushed with `git log --oneline origin/main..HEAD` rather than inferring it from a file list.

**Mutation scope is exactly `src/domain/**` + `src/use-cases/**`, with only two structural exclusions:**

1. `**/*.test.ts`: test files have no logic to mutate
2. `**/ports/**`: port files are type-only declarations (zero runtime, zero mutants)

**No file gets a per-file exclusion just because its tests feel awkward.** If a file produces equivalent or timing-flaky mutants, the right answer is one of:

1. Tighten the test (assert the specific behaviour the mutant breaks).
2. Refactor the production code to be more directly testable (extract pure helpers from a dispatch loop, etc.).
3. Improve fixtures so timing isn't load-bearing.

Skip lists rot: the next person assumes a file was untestable when really it was just inconvenient that day. If you're tempted to exclude a file, that's a smell: fix the test instead.

ESLint must ignore `.stryker-tmp/` and `reports/` so Stryker scratch dirs do not get linted (see `references/bun-typescript.md`).

### Commit message format (commit-msg hook)

The fast gates above run on the **staged diff** via `pre-commit`. Commit *messages* are validated by a separate git hook, `commit-msg`, which fires after you write the message, `assets/commit-msg`, installed to `.githooks/commit-msg` and picked up by the same `core.hooksPath`. It is a different hook on a different input (SKILL.md hard rule 23).

The contract is [Conventional Commits](https://www.conventionalcommits.org):

```
type(optional-scope)!: subject
```

- **type**: one of `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, `revert` (the `@commitlint/config-conventional` set).
- **scope**: optional, lowercase, in parentheses: `feat(auth):`, `fix(api):`.
- **!**: optional, marks a breaking change: `refactor(orders)!:`.
- **subject**: required, no trailing period, header ≤100 chars.

Why a hook and not just a guideline: the commit log is the project's changelog and `git bisect` surface. A machine-readable `type`/`scope` lets tooling derive release notes, group history, and flag breaking changes. A soft "please use Conventional Commits" drifts within a week; the hook keeps every commit on `main` honest and rejects `wip:`, `update stuff`, `Fix: thing`, and the like. git-generated `Merge`/`Revert`/`fixup!`/`squash!` headers are passed through untouched.

The shipped `assets/commit-msg` is a **dependency-free shell validator**: it matches the hand-rolled style of the other gate scripts and adds nothing to `package.json`. The Next.js monorepo variant enforces the identical grammar through `@commitlint/config-conventional` (already in its root toolchain) wired as a `simple-git-hooks` `commit-msg` step; see `references/nextjs-monorepo.md`. Either way the grammar is the same, only the validator differs.

**The hook is the first line, CI is the one that cannot be skipped.** `git commit --no-verify`
walks straight past the hook, so `assets/check-commit-messages.sh` re-runs the identical grammar
in CI over every commit in the pushed range: the PR's base branch on a pull request; on a push,
`github.event.before..HEAD`, which the shipped workflow exports as `GITHUB_EVENT_BEFORE` (on a
push to `main` the checkout has `HEAD == origin/main`, so a range against `origin/main` is empty
by construction and would pass without checking anything); `HEAD~1..HEAD` when neither resolves;
never the whole history, so a repo adopting the standard is not failed on commits nobody can
rewrite. `check-commit-range.sh` resolves its base the same way. The message gate delegates to the
hook script rather than restating the pattern, so local and CI cannot drift apart. Merge commits
are excluded; the hook's own Revert / fixup! / squash! / amend! exemptions still apply.

For Husky, copy `assets/commit-msg`'s body into `.husky/commit-msg`.

### Periodic audit: surface dead code in `test-helpers`

`src/test-helpers/**` is in the normal coverage skip list because it is test infrastructure, not production code. But that means **dead helpers can sit there at <100% indefinitely**. The fix is a periodic audit:

Once per release (or quarterly), temporarily remove the `test-helpers` skip from `scripts/check-coverage.ts` and run `bun run coverage`. Anything below 100% is one of two things:

1. **Dead code.** Delete it. Coverage gaps are a YAGNI smell-detector. (Real example: `networkThrow(message)` in `fetch-mock.ts` was a speculative helper that no test ever called: every test inlined `respond: () => { throw new TypeError(...) }` instead. Deleted.)
2. **Untested defensive code** (e.g. `installFetchMock`'s "no handler matched" guard). Add a one-test smoke block: they're load-bearing even when normal tests don't hit them.

Restore the skip after the audit. Schedule it on a calendar; the longer between audits, the more dead code accumulates.

### Discipline tripwires (rules 26-30)

Five shipped guards move the mechanical slices of the production disciplines into the machine tier. Each checks the **staged diff** (like the staged secret scan), so it blocks a violation entering history without flooding a brownfield tree; each takes `--all` for a tree-wide adopt-mode audit; exceptions ride on path conventions, never inline suppressions (rule 15).

| Guard | Rule | Blocks |
|:---|:--:|:---|
| `assets/check-identity.sh` | 26 | a person, employer, or client named in file contents: the committer's multi-word name (both orders) or email, the history's authors under `--all`, `IDENTITY_DENYLIST` entries; handles, noreply addresses, CODEOWNERS and `.mailmap` pass. A core gate, run by every shipped hook and CI workflow; it blocks in a public repository and warns in a private one (`.github/visibility`, which CI checks against `REPO_VISIBILITY`) |
| `assets/check-pii-channels.sh` (core, via `check-disciplines.sh`) | 27 | a natural identifier (thirteen names, any casing or prefix) in a query string (literal or via `new URLSearchParams`), a logger message interpolation (the call joined with up to 3 following lines), a Java `@QueryParam` |
| `assets/check-io-deadlines.sh` (core, via `check-disciplines.sh`) | 29 | an infra `fetch` / `globalThis.fetch` call with no `AbortSignal.timeout(` / `signal:` within the 8 lines after it, comments stripped (Java `HttpClient` with no `.timeout(` / `connectTimeout` in the file) |
| `assets/check-data-lifecycle.sh` (core, via `check-disciplines.sh`) | 30 | a hard delete in app code (erasure/retention/prune/sweep paths exempt, matched on the path only); DROP COLUMN / DROP TABLE / RENAME / TRUNCATE / ALTER COLUMN TYPE outside a `*contract*` migration |
| `assets/check-isolation-tests.sh` (opt-in, where tenants exist) | 28 | a new route file with no test named for it that asserts 404 inside a test block, comments excluded (`*public*`/`*health*`/`*to-response*` exempt) |

Since 2026-09-10 four of the five are core gates: the identity guard and, through `assets/check-disciplines.sh` (one hook step, one CI step, every guard run even after one fails), the personal-data, deadline and data-lifecycle guards, each inert in a repo without the concern and wrong in any repo with it. The isolation guard is the exception and stays opt-in **where tenants or owners exist**: it demands a cross-tenant 404 test of every new route, which is wrong for a single-user app; wire it beside the wrapper in the hook and CI when the concern exists. They are tripwires, not proofs; the discipline references keep the full review duty. The repo smoke test exercises all four so a regression in a guard fails CI here first.

### Never bypass with `--no-verify`

`git commit --no-verify` skips every gate. It is reserved for genuine big-bang changes (initial scaffolds, mass-rename refactors, generated-file updates), never for a failing check. **Justify every bypass in the commit body.** Do not normalise bypassing.

If a check is wrong for the codebase, fix it at the project level: raise or lower a rule's severity in `eslint.config.js`, adjust a coverage gate in `scripts/check-coverage.ts`, refactor a flaky test, and commit the fix. Same discipline as the no-inline-ignore rule: refactor or reconfigure, never suppress.

### Adapt for Husky or another hook manager

If the repo already uses Husky, drop the body of `assets/pre-commit` (from `set -euo pipefail` onwards) into `.husky/pre-commit`. The shebang and the `git config core.hooksPath` step are unnecessary; Husky handles them.

## Dependency CVE scanning (CI)

Gate 2 pins every dependency to a concrete version for supply-chain safety, but a pinned `^1.2.3` can still *be* a known-vulnerable version: pinning stops silent upgrades, it does not scan. The only dependency scanner the toolchain ships (the Snyk IDE extension, see `references/bun-typescript.md`) runs IDE-side, so CI and the pre-commit hook never see its findings, the same drift problem that motivated mirroring SonarLint into ESLint (see *SonarLint findings caught at lint time* above). The fix is a deterministic CVE scan in CI.

**Tool: `bun audit`.** Bun-native, no new dependency, reads the resolved tree from `bun.lock`, and exits non-zero when it lists a vulnerability. `--audit-level=high` filters to high/critical; an unfixable advisory is allow-listed with `--ignore <id>` **at the workflow level, with a reason**, never inline, the same rule as a project-level ESLint severity change.

**It is a CI job, not a hook gate.** CVE feeds change daily, independent of your diff. Blocking a 10-file commit because a new advisory dropped overnight in an *untouched* dependency fails in the wrong place. So the scan runs in CI on two triggers, each doing a different job:

- **A scheduled daily run** is the real watchdog: it is the only thing that catches a newly-disclosed CVE in a dependency *nobody touched*. A red scheduled run is the signal; wire it to an issue or chat alert if you want (out of scope here).
- **A pull-request run scoped to the manifests and `bun.lock`** blocks vulnerabilities a PR *deliberately introduces*, while never red-flagging PRs that don't change dependencies.

`.github/workflows/audit.yml` ships as `assets/audit.yml` (the Bootstrap checklist copies it; one copy, read it there): a daily schedule plus pull requests that touch any `package.json`, `bun.lock` or a vendored copy of the standard; `bun install --frozen-lockfile` (which also fails on lockfile drift, a supply-chain check in itself), `bun audit --audit-level=high`, then `check-skill-pin.sh` (`references/governance.md`); read-only permissions. An advisory with no upstream fix is allow-listed in that file, project-level, with a reason: `bun audit --audit-level=high --ignore GHSA-xxxx-xxxx-xxxx` beside a comment naming the date and the tracking link.

`--audit-level=high` fails the job only on high/critical advisories; moderate and low are reported but do not block: run `bun audit` locally to see the full list. The scan covers **all** dependencies, not `--prod` only: dev and build tooling are part of the supply-chain attack surface CI exists to watch.

Beyond this scan, CI gains a job whenever the matching concern exists in the repo: an eval gate on any LLM hole (`references/ai.md`), an axe scan on a UI (`references/product.md`), a load-test threshold on a hot route (`references/reliability.md`), the compose portability boot, deployment events, and the scheduled restore drill (`references/delivery.md`). Each is prescribed by its reference; the wiring is per-repo.

## Verification discipline (a control is a hypothesis until tested)

The gates make the standard executable; this section is about not trusting a gate, a guard, or a fix until something has tried to defeat it.

- **Test the bypass, not the happy path.** A guard proves nothing until a test walks the forbidden path and is refused: wrong role 403, missing token 401, cross-tenant 404, forged trust header inert. See `references/testing.md` (Bypass tests) and `references/isolation.md`.
- **Audit the seams between systems.** The dangerous gap lives where two individually-correct systems meet (proxy to app, edge to service); test the path a real request travels, not each box in isolation.
- **Fix the class, not the instance.** When a flaw is found, assume it repeats wherever the pattern does: enumerate with `rg` first, fix every hit, then add a CI guard that fails if the pattern returns.

```bash
rg -n '(db\.query|db\.execute)\(`.*\$\{' -- 'src/**/*.ts'   # enumerate the whole class
# then a CI step: if rg -q <same pattern>; then echo "interpolated SQL sink" >&2; exit 1; fi
```

- **Compliance is not proof.** A ticked checklist and a passed audit describe paperwork. The standard is a runnable check: "show me how you verify it, and let me run it myself." Evidence is the exit code of a committed script anyone accountable can execute, never a screenshot of a green run (`references/governance.md`, owner-verifiable done).
- **Generated code meets the same bar (provenance is not proof).** Code from a scaffolder, a generator, or an AI assistant runs through the identical hooks, gates, suite, and review a human's would; the reviewer reads the diff, not the attribution. No `--no-verify` because "the tool wrote it".
- **Prefer failing loud.** A gate that stays green for the wrong reason lies: that is why untested files enter coverage at 0% (the preload), why the mutation gate exists at all, and why each new gate should be tried against a known violation once before it is trusted (the smoke tests do exactly this for the shipped configs).
- **A skill description is a triggering contract; edits to it rerun the trigger eval.** Any change to a `SKILL.md` frontmatter description runs its eval set before landing (in the atelier skill repository itself: `bash scripts/trigger-eval/run.sh <set> <skill-dir>`, with the `suite-routing.json` set and `TRIGGER_EVAL_SUITE` when wording could shift which suite skill wins a query; a repo that ships its own skills keeps an equivalent eval set). A description tuned by feel regresses silently; the eval is one command.

## README consistency

The README is the contract with anyone who clones the repo. If it lies, the change is broken even if the tests are green. Audit it twice: once before declaring a task done, and once more before ending the session.

### When to audit

Run the audit whenever the working tree has uncommitted changes on a non-`README.md` file. Skip it only when the changes are clearly internal-only: private helpers, test-only refactors, formatting passes, dependency bumps that do not change usage.

### What to walk

The user-visible surface area is the set of facts the README documents about the project from the outside. For an atelier-shaped repo, that is roughly:

| Surface | What changed in the session that would invalidate the README |
|:---|:---|
| Install / setup steps | Added a system dep (`gitleaks`, `bun`), changed the install command, moved a config file the install copies |
| `package.json` scripts | Added/renamed/removed any of `test`, `lint`, `typecheck`, `coverage`, `mutate:*`, etc.; changed what one of them does |
| CLI flags / subcommands | New flag, renamed flag, changed default, removed flag: both the flag itself and the example invocations in the README |
| Env vars / config files | New `process.env.X` read in `src/composition/env.ts`; new entry in `.env.example`; new key in `bunfig.toml` |
| Top-level layout / architecture diagram | New top-level folder, renamed folder, deleted folder: the README's tree diagram and any prose that names paths |
| Public exports | A function/type/module the README documents as the API surface (not the same as "everything exported from `src/`") |
| Pinned versions | The README mentions "Bun ≥ X" or "Next.js Y" and the actual `package.json` / `bunfig.toml` pin moved |

If the audit finds drift, fix the README in the **same commit** as the code change: drifted READMEs across separate commits are how docs rot.

### Past breakages this rule catches

- CLI flags renamed but `--flow` examples stayed
- Scripts added to `package.json` but not listed
- Folders deleted (or renamed) but the architecture diagram still referenced them
- Coverage and prompt sections missing entirely from a change set that introduced them
- Install one-liner stayed pointing at a deprecated tool while the docs body listed the new one

### Five-check task-done gate

The four inner-loop checks from the top of this file (`bun run test`, `bun run lint`, `bun run typecheck`, `bun run coverage`) plus a fifth: `README.md` audited against the surface table above, either updated, or a one-sentence "nothing user-visible changed".

### End-of-session re-audit

A session usually contains several back-to-back tasks. Each one might pass its task-done audit, then the next one drifts the README again. So re-walk the surface table once more before stopping the session, even if every individual task said "nothing user-visible changed", the cumulative diff often does. State the result in one sentence: "README still current" or "README updated for X, Y, Z".

## Editor configuration that keeps formatting stable

Two guardrails prevent Prettier ↔ VS Code TS-formatter drift:

1. `source.fixAll.eslint` on save applies the ESLint-with-Prettier rules **after** whatever formatter handled the file, so the lint rules always have the last word. TS/TSX files format with `vscode.typescript-language-features` (its output is then normalised by the ESLint fix pass); everything else defaults to `dbaeumer.vscode-eslint`. The canonical per-variant `.vscode/settings.json` blocks live in `references/bun-typescript.md` and `references/nextjs-monorepo.md`.
2. The pre-commit hook catches staged-lint and typecheck drift at commit time; the full suite, coverage, and mutation catch the rest in CI.

```json
// .vscode/settings.json (excerpt: full blocks in the variant references)
{
  "editor.formatOnSave": true,
  "editor.defaultFormatter": "dbaeumer.vscode-eslint",
  "editor.codeActionsOnSave": {
    "source.fixAll.eslint": "explicit"
  },
  "[typescript]": { "editor.defaultFormatter": "vscode.typescript-language-features" },
  "[typescriptreact]": { "editor.defaultFormatter": "vscode.typescript-language-features" }
}
```

## TypeScript config for VS Code + Bun interop

`bun run typecheck` (invoking `tsc --noEmit`) finds the `bun:test` module via type-acquisition heuristics. VS Code's TypeScript server does not, and errors `Cannot find module 'bun:test'`. Fix with an explicit `"types"` array in `tsconfig.json`:

```jsonc
{
  "compilerOptions": {
    "types": ["bun"]
    // ... rest of config
  }
}
```

After the change, restart the TS server in VS Code (Cmd/Ctrl + Shift + P → "TypeScript: Restart TS Server"). The CLI typecheck passes either way; the editor needs the explicit list.

## Summary

- **Inner-loop checks, always, in order:** `bun run test`, `bun run lint`, `bun run typecheck`, `bun run coverage`.
- **Zero warnings, zero inline ignores.** Refactor or change severity at the project level; never suppress per-line.
- **Coverage gates per-tier:** 100% on `domain` + `use-cases`, 80% on `composition` + `infra` + `presenter`, skip `test-helpers` and `main.ts` only. `build-deps.ts` is now in scope (testable via optional config DI).
- **SonarLint parity at lint time** via `eslint-plugin-sonarjs` + type-aware `@typescript-eslint` rules.
- **Pre-commit hook runs the fast gates** (commit size, package.json, the staged secret scan, identity, the discipline tripwires, lint:staged, typecheck); **CI (`assets/ci.yml`) runs the full set** and is the required merge check: commit messages over the pushed range, strict lint, typecheck, the whole test suite, coverage, and mutation on the changed files, on a frozen lockfile; the full mutation sweep (`assets/mutation.yml`) and the CVE scan (`assets/audit.yml`) are their own scheduled workflows.
- **Commit identity** (rule 26): contributor identity in commit metadata is normal and never a finding; file contents never name a person, an employer, or a client. Scrubbing a mention from pushed history takes a gated `git filter-repo` rewrite plus a force-push, and the host may keep the old commits cached.
- **Dependency CVE scanning lives in CI, not the gate** (`bun audit --audit-level=high`): a daily scheduled watchdog for new CVEs in untouched deps, plus a PR run scoped to `package.json` / `bun.lock` for deliberately-introduced ones.
- **Mutation testing on the changed files** (Stryker, ≥90% break threshold, `mutate:changed` in CI on every event, the full sweep daily) makes "tests don't actually pin behaviour" findable.
- **Verification discipline:** a control is a hypothesis until a test walks the forbidden path; test the bypass, audit the seams, fix the class not the instance, and proof is a runnable check, not a checklist or a screenshot. Generated code meets the identical bar; provenance is not proof.
- **Commits stay small:** ≤10 files AND ≤300 lines per commit. The hook enforces it.
- **Periodic audits**: once per release, drop the `test-helpers` skip and run coverage; anything below 100% is dead code or untested defensive code.
- **README.md is part of the change set.** Re-read it before declaring any task done.

## Red flags the gates miss

SKILL.md keeps the rule; this is the symptom list to read when reviewing. Any hard-rule breach is a red flag by definition; these are the shapes the mechanical gates cannot see (the SonarJS S-rules and the `bunfig.toml` coverage-threshold trap have their own sections above).

- Untrusted input reaching a sensitive sink (SQL, shell, filesystem, HTTP, HTML, redirect) without a branded-type checkpoint between them.
- A secret (token, password, API key, PII) interpolated into a log line, or placed in a `NEXT_PUBLIC_*` env var.
- Creating, editing, deleting, renaming, or skipping a test file, or weakening an assertion, changing an expected value, or commenting out a case, without first showing the user and getting an explicit yes (rule 24). Weakening a failing test to go green instead of fixing the code is the worst of these; the default for a red test is to fix production code.
- Running `git commit` or `git push` without the user's explicit confirmation (rule 25). Staging and proposing the commit is the agent's role; pulling the trigger is the user's. "Do it" on a task is not commit approval: show the proposed commit and ask.
- An infra adapter exported with no test seam at all: no custom-fetch DI, no `createXFromApi(api: XApi)` factory, no sync-builder export (`references/testing-infra.md`). Without a seam, someone will reach for `mock.module` on the next test. Expose one from day one, even before the first test exists.
- Adding a new `src/infra/*.ts`, `src/composition/*.ts`, or `src/presenter/*.ts` file without regenerating `scripts/coverage-preload.ts` in the same commit (`bun run scripts/regenerate-coverage-preload.ts`). Untested infra files are invisible to `bun test --coverage` unless something imports them; the preload makes them appear at 0% so the gate can fail loudly.
- An inline suppression of any tool: `// eslint-disable*`, `// @ts-ignore`, `// @ts-expect-error`, `// snyk-ignore`, `// sonar-ignore`, `// deepcode ignore`, `// istanbul ignore`. Refactor, or change rule severity at the project level.
- `Record<K, V>` when the key set is open. Use `Partial<Record<K, V>>` so the type tells the truth about missing keys.
- Domain-specific data (brand lists, flow slugs, tier rates, tenant names) hardcoded as string-literal unions or records in framework code. Drive from env or config files; keep the framework generic.
- A per-file exclusion in `stryker.conf.json` for "the tests are awkward". Skip lists rot. The only structural exclusions are `**/*.test.ts` and `**/ports/**`. If a file produces equivalent or flaky mutants, tighten the test or refactor the production code, never add it to a skip list.
- A commit exceeding 10 files OR 300 lines (insertions + deletions) without a clear big-bang justification (initial scaffold, mass-rename, generated files). Split into smaller coherent slices. The pre-commit gate enforces this; do not normalise `--no-verify`.
- A commit message that is not Conventional Commits: no `type:` prefix, an unlisted type (`wip:`, `update:`), a capitalised type, a trailing period, or a >100-char header. The `commit-msg` hook rejects these (hard rule 23); write `type(scope): subject` the first time rather than reaching for `--no-verify`. A repo with the fast-gate `pre-commit` installed but no `commit-msg` hook is half-protected, wire both.
- A composition root or wiring file declared "untestable" and skipped. The two ergonomic switches make any composition file 100%-testable: parameterise every state-source (path, env var, clock) and inject every output sink (logger, sender). See `references/architecture.md` (Composition root testability).
- An assignment to `process.env.X = ...` anywhere outside `*.test.ts` (and even there, only inside `beforeAll`/`afterAll` with a saved-and-restored original). `process.env` is shared mutable state: pass values as parameters instead (`references/security.md`; config is read once, in `src/composition/env.ts`).
- A rule 21–22 breach anywhere in the UI: a hook call inside `src/components/**`; an import of `src/lib/**`, `src/config/**`, or `next/*` in a design-system component; `'use client'`, translation resolution, `process.env`, or data fetching in one; a downward import (an atom importing a molecule); a Tailwind utility string outside `src/components/**` (tokens in `globals.css` aside); or free-form `className`/`style` in a molecule/organism public API. State is hoisted, links/images arrive as injected `ComponentType` props, and visual variation is a typed variant prop.
- Personal data in the wrong channel (rule 27): an email, a name, or user-typed text in a log line, a URL, a query string, or a third-party analytics event; a new loggable field added without checking the redaction keys.
- An isolation breach in the making (rule 28): an owner id read from a URL, header, or body; a query that returns unscoped rows when the owner is missing; an owner-scoped endpoint landing without its cross-tenant 404 test; a sequential integer id in a public URL.
- An outbound call with no timeout, an unbounded or unjittered retry loop, or a retried non-idempotent operation without an idempotency key (rule 29). A fire-and-forget side effect after a commit that should be an outbox row.
- A hard DELETE on live data, a hand-run or destructive in-place schema change, a shipped contract field renamed or dropped in one step (rule 30), or a mutable shared record written back without its version check (rule 31).
- An LLM breach (rule 32): a provider SDK imported outside one infra adapter; a floating model alias (`latest`, an undated name); model output consumed without a schema checkpoint; a model-requested action executed without server-side authorization for the actual caller; a prompt or pin change shipped without its eval case set and threshold runner in the diff; a metered AI route without a per-caller spend gate.
- Hand-rolled session tokens, password hashing, or crypto; an admin surface without SSO plus MFA; a route that skipped the authenticated-by-default baseline without a written exception (rule 33).
- A production dump restored into a lower environment, or a fixture carrying a real person's data (rule 34).
- User-facing failure with no designed state: a raw status code or stack trace shown to a person, copy hardcoded outside the i18n catalog, a clickable `div` where a `button` belongs, or a flow that cannot be completed by keyboard (`references/product.md`).
