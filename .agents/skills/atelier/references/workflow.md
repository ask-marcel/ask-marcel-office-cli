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

**PLAN.md is not LESSONS.md.** `PLAN.md` is the *mutable current plan* and is rewritten and overwritten freely. `.claude/LESSONS.md` is *append-only memory* (decisions, gotchas) and is never rewritten. A durable decision that outlives the task graduates from a PLAN breadcrumb into a `[decision]` lesson; the plan step itself is transient. See `references/lessons.md`.

**On resume.** Start of session, read `.claude/PLAN.md` (alongside the lesson files). If it shows an unfinished task, continue from the first unchecked step rather than re-planning. If the user's new request supersedes the open plan, say so in one sentence and overwrite it.

**Within a long run.** The live plan is also your context-budget checkpoint, not only a crash-recovery file: keeping it current means a long agentic run degrades gracefully instead of hitting a context limit blind, because the next step and its DoD are always on disk. When a task is too large to finish in one context window, decompose it into independently-checkpointed steps (or subagents) rather than driving one context past the wall.

## The four-check loop (after every change)

```bash
bun test           # must pass
bun run lint       # 0 errors AND 0 warnings
bun run typecheck  # tsc --noEmit, clean
bun run coverage   # per-directory thresholds pass
```

If any of the four fail, fix the cause and re-run all four. Do not move on while one is red. Warnings have repeatedly hidden real issues (silent precedence bugs, dead returns, suppressed scanners); the zero-warning rule is not cosmetic.

`package.json` scripts:

```json
{
  "scripts": {
    "lint": "eslint --cache --max-warnings=0",
    "lint:strict": "LINT_STRICT=1 eslint --max-warnings=0",
    "typecheck": "tsc --noEmit",
    "coverage": "bun run scripts/check-coverage.ts"
  }
}
```

`bun run lint:strict` (~25 s) sets the env var `LINT_STRICT=1`; the same `eslint.config.js` reads `process.env['LINT_STRICT']` and conditionally adds a type-aware block (`parserOptions.projectService: true` plus `@typescript-eslint/no-unnecessary-type-assertion` and `@typescript-eslint/prefer-promise-reject-errors`). One config file, two modes, no separate `eslint.strict.config.js` to keep in sync. CI runs the strict version as a merge gate; the pre-commit hook runs the fast, non-type-aware `lint:staged` on the staged files, so run `bun run lint:strict` yourself in the inner loop to see the type-aware findings before you push.

## Zero warnings; no inline ignores

`bun run lint` is considered failing if it reports any warnings, not just errors. Two acceptable ways to clear a finding:

1. **Refactor the code** so the rule no longer fires. This is the default. If Snyk's string-literal-adjacent-to-key heuristic flags `const apiKey = 'sk-...'`, build the value at runtime from env vars. If `no-await-in-loop` fires, restructure the loop into `Promise.all` or accept the sequential cost with a targeted severity change.
2. **Configure rule severity at the project level** in `eslint.config.js`, with a comment explaining why. Reserved for rules that produce only false positives in this codebase's idioms — branded types, string-literal unions, bounded regexes, `security/detect-object-injection` on `Record<K, V>` lookups. For security-plugin rules, the two-part test below must be satisfied before disabling. Atelier already does this with `unicorn/no-null`, `unicorn/empty-brace-spaces`, `security/detect-object-injection`, `security/detect-unsafe-regex`, and `security/detect-non-literal-fs-filename`.

**Forbidden everywhere, no exceptions:**

- `// eslint-disable`, `// eslint-disable-next-line`, `// eslint-disable-line`
- `// @ts-ignore`, `// @ts-expect-error`
- `// deepcode ignore`, `// snyk-ignore`, `// sonar-ignore`, `// istanbul ignore`
- Any equivalent from another tool

If a rule needs suppression for a single line, the suppression is a lie — either the rule is wrong for this codebase (change severity at the project level) or the code is wrong for this codebase (refactor). A per-line suppression hides both.

### Project-level rule disabling: a two-part test

Before adding a rule to the `'off'` list, both of these must be true:

1. **Every fire in this codebase is a false positive for our idioms.** Not "most", not "the current ones". Branded types + `Record<K, V>` lookups, bounded regexes with documented inputs, `chmodSync(mkdtempSync(...))` inside FS-adapter tests — these never represent a real exploit in the atelier style. If even one fire out of ten is a genuine finding, leave the rule on and refactor the other nine.
2. **The pattern the rule would catch if it fired correctly is something the production code cannot produce.** `security/detect-non-literal-fs-filename` matters for `node:fs` calls in production; atelier production code uses `Bun.file` instead (which the rule does not watch), so disabling globally loses nothing on the real attack surface. If production *could* produce the pattern, disabling masks a real vulnerability class.

Never disable a rule globally just to silence a single test, a single commit, or a single file. If the fire is localised, the right tool is a narrower `files:` scope in the ESLint config (disable the rule for `**/*.test.ts` only, for example) — still at the project level, still with a comment, never inline.

## Keeping `coverage-preload.ts` in sync (auto-regeneration)

`scripts/coverage-preload.ts` lists every file in `src/{infra,composition,presenter}/` so the coverage table can include them at 0% if untested. Keeping this file in sync by hand is tedious and the failure mode is silent — a missing import means the new file never appears in the coverage report and the gate trivially passes.

One script handles this, shipped in the skill at `assets/`:

| Script | Job |
|---|---|
| `assets/regenerate-coverage-preload.ts` | Walks `src/{infra,composition,presenter}/`, excludes `*.test.ts` / `ports/` / `index.ts`, writes a fresh `scripts/coverage-preload.ts` |

`scripts/coverage-preload.ts` itself is always **generated**, never hand-written — copy the regenerate script into a new repo and run it once to create the initial preload.

**Two modes:**

```bash
# Write a fresh preload from the current src/ tree.
bun run scripts/regenerate-coverage-preload.ts

# Exit non-zero if the on-disk file is out of sync — for pre-commit / CI.
bun run scripts/regenerate-coverage-preload.ts --check
```

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
- `google.drive({ version: 'v3', auth })` — instantiates a real Google client; can't be exercised without real credentials
- `new MongoClient(url)` — opens a real driver
- A `process.on('SIGTERM', ...)` handler that the test runner never sends

These lines are covered by the **production-wiring smoke test** described in `references/testing-infra.md` whenever possible — the smoke test calls `createX(realDeps)` with a placeholder, which exercises the wiring line and asserts the resulting port has the right method shape.

When even the smoke test cannot reach a line (a closure inside a method that requires real SDK behaviour to enter), accept it as exempt. The 80% gate on `src/infra/**` is calibrated for this; a single adapter at 88-95% line coverage with the rest of the file fully tested is healthy.

**Rule of thumb.** Lines that exist solely to bridge into a third-party SDK (dynamic imports inside a closure, real-factory pass-throughs, SDK-instantiation one-liners) are exempt from line-coverage when the file's other paths bring it above the per-tier gate. Do not lower the gate; do not add a per-file skip in `bunfig.toml`. Just accept that the bridge line is the cost of doing business with the SDK and the gate is permissive enough to absorb it.

If a single file is dragged below the 80% gate by SDK-bridge lines alone, the right move is usually to refactor — split the bridge into a thinner `createX(realSdk)` that only does the instantiation, and a fatter `createXFromApi(api)` that holds all the logic. Then the bridge file is one or two lines (still uncovered, but tiny) and the logic file is fully tested.

## Coverage gates (per-tier, enforced by custom script)

Bun's built-in `coverageThreshold` is a single global number. It cannot express "100% on the domain, 80% on infra, skipped on test-helpers". The repo enforces per-tier rules via `scripts/check-coverage.ts`, which runs `bun test --coverage`, parses the text report, and applies path-prefix rules.

A ready-to-copy `check-coverage.ts` lives in the skill at `assets/check-coverage.ts`. It exposes `COVERAGE_RULES` and `SKIPPED` as top-of-file constants so tuning per-project takes a one-line edit.

| Path | Threshold (functions & lines) |
|:---|:---|
| `src/domain/**` | **100%** |
| `src/use-cases/**` (including `ports/`) | **100%** |
| `src/composition/**` (env.ts AND build-deps.ts) | **80%** |
| `src/presenter/**` | 80% |
| `src/infra/**` | **80% from day one** — not "once tests exist" |
| `src/test-helpers/**` | skip during normal runs (audit periodically — see below) |
| `src/main.ts` | skip (entry point; verified by integration) |

**`build-deps.ts` is no longer skipped.** The earlier policy excluded it as "composition root, verified live, no logic worth unit-testing". That was hedging. The composition root becomes fully unit-testable when (a) every "where do I read state from" point (file path, env var, system clock) is parameterisable, and (b) every "what do I write to / log to" sink can be injected as a port. See `references/architecture.md` (Composition root testability) for the optional-config-DI pattern.

Every `src/infra/*.ts`, `src/composition/env.ts`, and `src/presenter/cli.ts` carries a real 80% gate — most end up at 100% once the three infra-test patterns (see `references/testing.md`) are in routine use. The "we'll add infra tests later" road leads to a coverage gate that trivially passes.

`bun run coverage` exits non-zero if any file falls below its gate and prints the offending paths with current-vs-required numbers. A tier summary at the end highlights the worst funcs/lines per tier, so a single sloppy file is visible without scrolling the per-file table.

If a file cannot hit the gate, the fix is usually **restructure the code so the dead branch goes away**, not lower the threshold. A threshold reduction must be justified in the commit message.

### The coverage preload (mandatory)

`bun test --coverage` only reports rows for files the test runner imports. Untested infra files — no `*.test.ts`, no test imports them — are silently absent from the table, which makes the per-file gate trivially pass. Every adapter you forgot to test becomes invisible instead of failing loudly.

The fix is a preload file that side-effect-imports every infra, composition, and presenter module, so they appear in the coverage table at 0% if no test exercises them. Generate it with `assets/regenerate-coverage-preload.ts` (§ Keeping `coverage-preload.ts` in sync, above) — never hand-write it. The generated file looks like:

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

**Wire it at coverage time only — NOT in `bunfig.toml`.** `scripts/check-coverage.ts` spawns:

```bash
bun test --coverage --preload ./scripts/coverage-preload.ts
```

Do not put `preload = [...]` under `[test]` in `bunfig.toml`. The preload pulls in heavy runtime deps (e.g. `googleapis`, `winston`, `twitter-api-v2`, `@ai-sdk/google`) that would add 1–2 seconds to every plain `bun test`. Loading it only at coverage time keeps the inner-loop fast without losing the gate.

**Maintenance rule:** every new file in `src/infra/`, `src/composition/`, or `src/presenter/` must be added to `coverage-preload.ts` in the same commit — run `bun run scripts/regenerate-coverage-preload.ts` and stage both files together. Enforcement is mechanical, not goodwill: wire `regenerate-coverage-preload.ts --check` as the pre-commit pre-flight (§ Keeping `coverage-preload.ts` in sync, above) so a forgotten regeneration blocks the commit.

### `bunfig.toml`: minimal, no `coverageThreshold`, no `preload`

The global `coverageThreshold` in `bunfig.toml` must be **absent** when the per-tier script owns enforcement. If set, Bun exits non-zero on the global threshold before the script runs, the script's first line (`if (result.status !== 0) return result.status`) bails, and the per-file violation breakdown never prints. The operator sees `error: script "coverage" exited with code 1` with no useful diagnostic.

`preload` must also be absent under `[test]` — see above.

Correct `bunfig.toml`:

```toml
[test]
coverage = true
coverageSkipTestFiles = true
coverageReporter = ["text"]
# NOTE: no `coverageThreshold` here — per-tier enforcement is in
# scripts/check-coverage.ts.
# NOTE: no `preload` here — the coverage preload is loaded only by
# `bun run coverage` (via --preload on the spawned bun-test command),
# so plain `bun test` runs stay fast.
```

When introducing a per-tier coverage gate in an existing repo, remove any global `coverageThreshold` *and* any `preload` from `bunfig.toml` in the same change.

## SonarLint findings caught at lint time

SonarLint runs IDE-side; CI and pre-commit do not see it. To keep IDE-only findings from drifting back in, ESLint is wired to catch them at lint time.

In `eslint.config.js` (one config, two modes — `LINT_STRICT=1` switches on the type-aware block). The **canonical, complete** flat config lives in `references/bun-typescript.md` (§ `eslint.config.js`) — the excerpt below shows only the SonarJS / type-aware slice this section is about, so if the two ever disagree, `bun-typescript.md` wins:

```js
import pluginJs from '@eslint/js';
import sonarjsPlugin from 'eslint-plugin-sonarjs';
import securityPlugin from 'eslint-plugin-security';
import tsPlugin from 'typescript-eslint';

export default [
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  securityPlugin.configs.recommended,
  {
    files: ['**/*.ts'],
    rules: {
      'no-restricted-imports': ['error', {
        paths: [{
          name: 'bun:test',
          importNames: ['mock'],
          message:
            '`mock` from bun:test is forbidden — it leaks across test files. Use dependency injection: refactor the production code to accept the SDK as a parameter, then pass a fake at construction.',
        }],
      }],
    },
  },
  // Type-aware rules — gated by LINT_STRICT=1. Inner-loop `bun run lint`
  // does not pay the ~25s parserOptions.projectService cost.
  ...(process.env['LINT_STRICT']
    ? [{
        files: ['src/**/*.ts'],
        languageOptions: {
          parserOptions: {
            projectService: true,
            tsconfigRootDir: import.meta.dirname,
          },
        },
        rules: {
          '@typescript-eslint/no-unnecessary-type-assertion': 'error',
          '@typescript-eslint/prefer-promise-reject-errors': 'error',
        },
      }]
    : []),
  sonarjsPlugin.configs.recommended,
  {
    // SonarJS rule overrides — always-on, justified per rule.
    rules: {
      'sonarjs/no-unused-vars': 'off',          // duplicates @typescript-eslint/no-unused-vars
      'sonarjs/no-empty-test-file': 'off',      // false positives on `describe` test layout
      'sonarjs/cognitive-complexity': 'off',    // function-size cap already covers this
    },
  },
  {
    rules: {
      'security/detect-object-injection': 'off',
      'security/detect-unsafe-regex': 'off',
      // false-positive on chmodSync(mkdtempSync(...)) in FS-adapter tests;
      // production uses Bun.file (not covered by this rule), so no real loss.
      'security/detect-non-literal-fs-filename': 'off',
    },
  },
];
```

The conditional block runs only when `process.env['LINT_STRICT']` is set, so the inner-loop `bun run lint` skips it entirely. `bun run lint:strict` is just `LINT_STRICT=1 eslint`.

### Common SonarJS findings and how to fix them

| Sonar ID | Symptom | Fix (never suppress) |
|:---|:---|:---|
| **S4325** | `x!` non-null assertion, or `x as Type` without real narrowing | Replace with a guard clause: `const found = xs.find(...); if (!found) throw new Error(...); return found;` |
| **S6594** | `"abc".match(re)` used for captured groups | Use `re.exec("abc")` — more efficient, avoids the global-flag trap |
| **S4123** | `await` on a matcher chain that is not a real `Thenable`, e.g. `await expect(p).rejects.toThrow()` | Use the `captureRejection(promise)` helper — see `references/result-type.md` |
| **S6551** | `String(err)` in a catch block | Use the shared `formatError(err: unknown): string` from `src/domain/utilities/format-error.ts` |
| **S6671** | `Promise.reject(value)` where value is not an `Error` | Change to `Promise.reject(new Error(...))`. For tests that deliberately reject with a non-Error, use a tiny `async (v: unknown) => { throw v }` helper |
| **sonarjs/void-use** | `void unusedParam;` to silence unused-var warnings | Drop the parameter entirely from the implementation. TypeScript's function-type **parameter contravariance** means a function with fewer parameters is assignable to a function-type with more. |

### Types must not lie

`Record<K, V>` says "every key maps to V". JavaScript's runtime says otherwise — missing keys return `undefined`. If the key set is open (user IDs, row IDs, environment variables), the honest type is `Partial<Record<K, V>>`.

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

Integrate to one branch — `main` (the trunk) — continuously. Commit straight to it, or through a branch that lives **less than a day** and merges back small. Long-lived feature branches are the thing this model exists to avoid: every day a branch diverges, the eventual merge gets riskier, review gets coarser, and `main` stops reflecting reality.

The rules already in this skill are precisely what makes committing to the trunk safe — trunk-based development is their reason for existing, not a separate concern:

- **Every commit keeps `main` releasable.** The fast pre-commit hook (below) blocks the obvious breakage locally (size, secrets, staged lint, types) in a few seconds, and CI holds the full line (the whole test suite, coverage, mutation, strict lint) as the required merge check, so what reaches the shared trunk is green.
- **Commits stay small** (gate 1: ≤10 files AND ≤300 lines). Small commits are the unit of continuous integration; they review in minutes, revert cleanly, and bisect precisely. A 300-line ceiling is a trunk-based ceiling.
- **History stays linear and legible** (Conventional Commits, `commit-msg` hook). A trunk read top-to-bottom is the changelog.
- **Incomplete work hides behind a flag, not a branch.** When a feature spans several commits, keep each commit green and the half-built path dark behind a feature flag or simply unreferenced — never park weeks of work on a divergent branch. This is the same instinct as YAGNI and "minimal": ship the smallest safe increment.

Practical loop: pull/rebase often to stay close to the trunk; run the four-check loop after every change; when green, propose the commit and commit it once the user confirms (SKILL.md hard rule 25 — the agent never commits or pushes on its own initiative); then push on their say-so. If a change is too big to land safely in one ≤300-line commit, split it into a sequence of green commits, not a long-lived branch. Releases are cut from the trunk (tag or release branch at the moment of release), never developed on for weeks beforehand.

This is the default for this codebase. It overrides any tooling habit of "branch first by default" — branch only when a short-lived branch genuinely helps (e.g. a PR-review gate your team requires), and merge it the same day.

## Commit identity (rule 26)

Every commit carries an author and a committer (each a name plus an email), taken from git config, and whatever they are becomes permanent public history the moment you push. Carrying the contributor's real identity in that metadata is normal, the default of the whole open-source world, and never a finding, an audit item, or a publish blocker.

File contents are the opposite. No tracked file ever names a person, an employer, or a client: not a name in a comment or a LICENSE holder line, not an employer's internal hostname in a config, not a client name in a fixture. A content mention outlives the commit that added it, travels with every copy and quote of the file, and once pushed cannot be removed by anything short of a history rewrite. Where a holder or author string is structurally required, use a neutral handle (e.g. `atelier`). Host control files whose format is identities (CODEOWNERS, `.mailmap`) are metadata in file form, not mentions; they are exempt. The cheap moment to catch a mention is review (atelier-review-me checks it); the expensive moment is after a push.

Secrets are the other real pre-publish concern: run `gitleaks detect` (the history-wide mode, not the pre-commit `protect --staged`) before the first push to a public host. Secrets in history are always findings; metadata identities never are.

**Scrubbing pushed history is a rewrite, gated and user-initiated.** A one-time, destructive operation; never run it unprompted (rule 25). Use `git filter-repo` (install: `brew install git-filter-repo`): `--replace-text` removes a mention from file contents across history, and `--mailmap` remaps commit metadata when the user wants that changed too:

```bash
# replacements.txt, one rule per line (mention ==> neutral replacement):
#   Old Name==>atelier
#   old@company.com==>atelier@users.noreply.github.com
git filter-repo --replace-text replacements.txt --force
git remote add origin <url>          # filter-repo strips the remote as a safety measure
git push --force-with-lease origin main
```

To also remap the author and committer fields, add `--mailmap` with `Intended Name <intended@email> <old@email>` lines. `filter-repo` rewrites every commit SHA, so this is a coordinated force-push: anyone holding a clone must re-clone.

**A force-push does not purge the old commits.** The rewritten branch no longer points at them, but the host keeps unreferenced commits reachable by their SHA, through cached views, and via any fork or open PR, until it garbage-collects on its own schedule. Treat a leaked commit as exposed even after the fix: rotate anything that was a live secret, and for a hard guarantee delete-and-recreate the repo or ask the host's support to purge.

## Gates: a fast pre-commit hook plus the full set in CI

The gate set has two homes, split by speed and not by importance (rule 15.1). The pre-commit hook runs the **fast gates** only, because a multi-minute hook trains `git commit --no-verify` (rule 15.3). Every gate, fast and slow, also runs in **CI**, the line that cannot be skipped and the required merge check (rule 4.6). The full test suite, per-tier coverage, and Stryker mutation are slow and grow with the codebase, so they live in CI and only in CI.

This is the **Bun-script variant's** mechanism. The Next.js monorepo uses `simple-git-hooks` (staged lint + typecheck + commitlint) instead, see `references/nextjs-monorepo.md`. Never install both: `core.hooksPath` and `simple-git-hooks` overwrite each other.

**The pre-commit hook (fast gates, target under ~5s):**

| # | Gate | Purpose | Typical time |
|:--:|:---|:---|:--:|
| 1 | `scripts/check-commit-size.sh` | <=10 files AND <=300 lines | <1s |
| 2 | `scripts/check-package-json.sh` | no `"latest"` / `"*"` / bare dist-tag | <1s |
| 3 | `gitleaks protect --staged` | secret scan on the staged diff | ~50ms |
| 4 | `bun run lint:staged` | ESLint on the staged TS files only | ~1-2s |
| 5 | `bun run typecheck` | `tsc --noEmit` clean | seconds |

Every gate here is O(staged files) or O(1). Typecheck is the one that grows with the whole codebase; if it exceeds the hook budget on your repo, move it to CI too. Never add the test suite, coverage, or mutation to the hook.

**CI (`assets/ci.yml`, the authoritative merge gate):** install on a frozen lockfile, then `check-package-json.sh`, `gitleaks detect` (full history), `lint:strict` (type-aware, zero warnings, ~25s), `typecheck`, `bun test` (the whole suite), `bun run coverage` (per-tier), `bun run mutate:changed` on a pull request or `bun run mutate` on main (1-3 min per file), and `bun audit`. Make it a required status check in branch protection (rule 13.2) so a bypassed hook is still caught.

A ready-to-copy hook lives in the skill at `assets/pre-commit`, the CI workflow at `assets/ci.yml`, and the staged-lint helper at `assets/lint-staged.sh`. The companion scripts (`check-commit-size.sh`, `check-package-json.sh`, `mutate-staged.sh`, `mutate-changed.sh`, `regenerate-coverage-preload.ts`) live alongside, plus `assets/commit-msg`, a separate git hook documented under *Commit message format* below.

### Install once per clone

```bash
mkdir -p .githooks scripts .github/workflows
cp <skill>/assets/pre-commit .githooks/pre-commit
cp <skill>/assets/commit-msg .githooks/commit-msg
cp <skill>/assets/ci.yml .github/workflows/ci.yml
cp <skill>/assets/lint-staged.sh scripts/lint-staged.sh
cp <skill>/assets/check-commit-size.sh scripts/check-commit-size.sh
cp <skill>/assets/check-package-json.sh scripts/check-package-json.sh
cp <skill>/assets/check-coverage.ts scripts/check-coverage.ts
cp <skill>/assets/regenerate-coverage-preload.ts scripts/regenerate-coverage-preload.ts
cp <skill>/assets/mutate-staged.sh scripts/mutate-staged.sh
cp <skill>/assets/mutate-changed.sh scripts/mutate-changed.sh
cp <skill>/assets/stryker.conf.json stryker.conf.json
chmod +x .githooks/pre-commit .githooks/commit-msg scripts/*.sh scripts/check-coverage.ts scripts/regenerate-coverage-preload.ts
git config core.hooksPath .githooks
# Generate the initial coverage-preload.ts from the current src/ tree
bun run scripts/regenerate-coverage-preload.ts
```

`core.hooksPath .githooks` picks up **both** `.githooks/pre-commit` (the fast gates, on the staged diff) and `.githooks/commit-msg` (Conventional Commits, on the message), one config, two hooks. `.github/workflows/ci.yml` is the authoritative gate set that runs every gate on every push and pull request.

Add to `package.json`:

```json
{
  "scripts": {
    "lint:staged": "bash scripts/lint-staged.sh",
    "mutate": "stryker run",
    "mutate:changed": "bash scripts/mutate-changed.sh",
    "mutate:staged": "bash scripts/mutate-staged.sh"
  },
  "devDependencies": {
    "@stryker-mutator/core": "^9.6.1"
  }
}
```

Install gitleaks (optional but recommended): `brew install gitleaks` on macOS, or grab a binary from `github.com/gitleaks/gitleaks/releases`. The hook degrades gracefully if `gitleaks` is missing — it warns and continues — so first-time clones don't break.

The `git config core.hooksPath .githooks` is the one step that is easy to forget. Without it, Git looks in `.git/hooks/` and your commit goes through unchecked. Document it in the repo's `README.md` install section.

### Commit size limits (gate 1)

`scripts/check-commit-size.sh` enforces the rule "≤10 files **AND** ≤300 lines (insertions + deletions)" — i.e. it blocks any commit that exceeds *either* threshold. The limits are conservative because they force the discipline; loosening them undermines the rule.

Why:
- Small commits are easier to review, revert, and bisect.
- Large commits hide bugs (one slip across 300 lines is hard to spot).
- Every commit on `main` becomes git history that the next engineer reads — keep each one a coherent slice.

When working on a feature, **commit as you go** — one focused slice at a time. The gate is the safety net, not the policy.

### Dependency hygiene (gate 2)

`scripts/check-package-json.sh` blocks any commit where `package.json` declares a version as `"latest"` or `"*"`. Every entry under `dependencies`, `devDependencies`, and `peerDependencies` must use a concrete version (`X.Y.Z`) or a real range (`^X.Y.Z`, `~X.Y.Z`, `>=X.Y.Z`).

Why:
- `"latest"` and `"*"` are non-deterministic. `bun install` on different days gives different `node_modules/` trees. The lockfile only partially mitigates this.
- The literal string `"latest"` semantically signals "always upgrade" — a silent-break footgun that can pull in a major version change between two checkouts of the same commit.
- You don't audit what you didn't expect to install. Hidden upgrades from `"latest"` are how supply-chain attacks land.

Workflow:

- **Adding a package.** `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev). Bun resolves the actual latest version at install time and writes it as `^X.Y.Z`. **Never hand-edit `package.json` to add a dep** — the gate may pass on a manually-typed `^1.2.3`, but you lose the auto-pinning convention and the muscle memory drifts.
- **Bumping every dep to current latest.** Run `bun update`. This rewrites the existing `^X.Y.Z` ranges to the latest matching versions and updates `bun.lock`. Commit both files in the same change. Do this on a deliberate cadence (start of a release, dependabot-style cron, etc.), not silently on every commit.
- **Bumping one specific dep.** `bun update <pkg>` for a constrained bump, or `bun add <pkg>@latest` to force the absolute current latest into the same `^X.Y.Z` slot. Either way, no `"latest"` ends up in the file.
- **Initial scaffold.** When using the skill's `package.json` skeleton (in `references/bun-typescript.md`), the version ranges are samples. Run `bun install` to resolve them, then `bun update` to pull each dep to its current latest, then commit both files together. Verify with `bash scripts/check-package-json.sh`.

The gate runs `grep -nE '"\*"|"latest"' package.json`. It catches the two bare strings only — version ranges like `^1.2.3` and `>=4.0.0` pass.

### Secret scanning with gitleaks (gate 3)

The hook runs `gitleaks protect --staged --redact --verbose --no-banner`. Two distinct gitleaks modes — pick the right one:

- **`gitleaks protect --staged`** — scans the staged-but-not-committed diff. Fast (~50 ms). Blocks re-introduction of secrets *before* they enter history. Use in pre-commit hooks.
- **`gitleaks detect`** — scans the entire git history (every commit, every file ever). Slow. Use for periodic audits or CI checks. **Does not** belong in a pre-commit hook.

Run `gitleaks detect` once before the first push to GitHub to catch anything that snuck in pre-hook.

### Mutation testing with Stryker (a CI gate)

[Stryker](https://stryker-mutator.io/) generates small "mutants" of the production code (e.g. `>` becomes `>=`, `&&` becomes `||`, `return x` becomes `return undefined`) and runs the test suite against each. A mutant that survives means your tests don't actually pin the behaviour they appear to.

The atelier policy: **every staged file under `src/domain/**` or `src/use-cases/**` must score ≥90% mutation score** before commit. The threshold is the `break` value in `stryker.conf.json`.

```jsonc
{
  "packageManager": "npm",
  "testRunner": "command",
  "commandRunner": { "command": "bun test" },
  "mutate": [
    "src/domain/**/*.ts",
    "src/use-cases/**/*.ts",
    "!**/*.test.ts",
    "!**/ports/**"
  ],
  "thresholds": { "high": 95, "low": 90, "break": 90 },
  "incremental": true,
  "incrementalFile": "reports/stryker-incremental.json",
  "concurrency": 4,
  "timeoutMS": 30000,
  "tempDirName": ".stryker-tmp",
  "cleanTempDir": true,
  "ignorePatterns": [
    ".claude/", ".agents/", ".githooks/", ".vscode/", ".git/",
    "docs/", "prompts/", "scripts/", "reports/", ".stryker-tmp/",
    "node_modules/", "*.md", "*.toml", "*.lock", "*.json"
  ]
}
```

There is no first-party `@stryker-mutator` Bun runner today (community plugins exist, but we don't depend on them), so we use the command runner — Stryker shells out to `bun test` once per mutant (~7 s on a typical codebase). `incremental: true` caches per-mutant results so unchanged code is not re-tested. `packageManager: "npm"` is needed because Stryker probes for a JS-ecosystem package manager and does not yet recognise Bun's lockfile. `ignorePatterns` skips non-source dirs from the sandbox copy — `.claude/` in particular often contains a symlink Stryker cannot copy (ENOTSUP).

**Three commands, three scopes:**

- **`bun run mutate`** — full run on `src/domain/**` + `src/use-cases/**`. Slow (1–2 hr on ~150 files). Periodic audit.
- **`bun run mutate:changed`** — files differing from `origin/main` plus uncommitted edits. Run during iteration. Override base ref with `BASE=HEAD~3 bun run mutate:changed`.
- **`bun run mutate:staged`**: files staged for the next commit. An optional local pre-push check; CI is the enforcing home, running `mutate:changed` on a pull request and `mutate` on main. Skips with exit 0 when no relevant files are staged, so commits to docs, tests, or scripts are unaffected.

**Mutation scope is exactly `src/domain/**` + `src/use-cases/**`, with only two structural exclusions:**

1. `**/*.test.ts` — test files have no logic to mutate
2. `**/ports/**` — port files are type-only declarations (zero runtime, zero mutants)

**No file gets a per-file exclusion just because its tests feel awkward.** If a file produces equivalent or timing-flaky mutants, the right answer is one of:

1. Tighten the test (assert the specific behaviour the mutant breaks).
2. Refactor the production code to be more directly testable (extract pure helpers from a dispatch loop, etc.).
3. Improve fixtures so timing isn't load-bearing.

Skip lists rot — the next person assumes a file was untestable when really it was just inconvenient that day. If you're tempted to exclude a file, that's a smell — fix the test instead.

ESLint must ignore `.stryker-tmp/` and `reports/` so Stryker scratch dirs do not get linted (see `references/bun-typescript.md`).

### Commit message format (commit-msg hook)

The fast gates above run on the **staged diff** via `pre-commit`. Commit *messages* are validated by a separate git hook, `commit-msg`, which fires after you write the message, `assets/commit-msg`, installed to `.githooks/commit-msg` and picked up by the same `core.hooksPath`. It is a different hook on a different input (SKILL.md hard rule 23).

The contract is [Conventional Commits](https://www.conventionalcommits.org):

```
type(optional-scope)!: subject
```

- **type** — one of `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, `revert` (the `@commitlint/config-conventional` set).
- **scope** — optional, lowercase, in parentheses: `feat(auth):`, `fix(api):`.
- **!** — optional, marks a breaking change: `refactor(orders)!:`.
- **subject** — required, no trailing period, header ≤100 chars.

Why a hook and not just a guideline: the commit log is the project's changelog and `git bisect` surface. A machine-readable `type`/`scope` lets tooling derive release notes, group history, and flag breaking changes. A soft "please use Conventional Commits" drifts within a week; the hook keeps every commit on `main` honest and rejects `wip:`, `update stuff`, `Fix: thing`, and the like. git-generated `Merge`/`Revert`/`fixup!`/`squash!` headers are passed through untouched.

The shipped `assets/commit-msg` is a **dependency-free shell validator** — it matches the hand-rolled style of the other gate scripts and adds nothing to `package.json`. The Next.js monorepo variant enforces the identical grammar through `@commitlint/config-conventional` (already in its root toolchain) wired as a `simple-git-hooks` `commit-msg` step; see `references/nextjs-monorepo.md`. Either way the grammar is the same — only the validator differs.

For Husky, copy `assets/commit-msg`'s body into `.husky/commit-msg`.

### Periodic audit: surface dead code in `test-helpers`

`src/test-helpers/**` is in the normal coverage skip list because it is test infrastructure, not production code. But that means **dead helpers can sit there at <100% indefinitely**. The fix is a periodic audit:

Once per release (or quarterly), temporarily remove the `test-helpers` skip from `scripts/check-coverage.ts` and run `bun run coverage`. Anything below 100% is one of two things:

1. **Dead code.** Delete it. Coverage gaps are a YAGNI smell-detector. (Real example: `networkThrow(message)` in `fetch-mock.ts` was a speculative helper that no test ever called — every test inlined `respond: () => { throw new TypeError(...) }` instead. Deleted.)
2. **Untested defensive code** (e.g. `installFetchMock`'s "no handler matched" guard). Add a one-test smoke block — they're load-bearing even when normal tests don't hit them.

Restore the skip after the audit. Schedule it on a calendar; the longer between audits, the more dead code accumulates.

### Discipline tripwires (rules 27-30, optional gates)

Four shipped guards move the mechanical slices of the production disciplines into the machine tier. Each checks the **staged diff** (like `gitleaks protect --staged`), so it blocks a violation entering history without flooding a brownfield tree; each takes `--all` for a tree-wide adopt-mode audit; exceptions ride on path conventions, never inline suppressions (rule 15).

| Guard | Rule | Blocks |
|:---|:--:|:---|
| `assets/check-pii-channels.sh` | 27 | a natural identifier in a query string (literal or via `new URLSearchParams`), a logger message interpolation, a Java `@QueryParam` |
| `assets/check-io-deadlines.sh` | 29 | an infra `fetch` / Java `HttpClient` with no deadline marker in the file |
| `assets/check-data-lifecycle.sh` | 30 | a hard delete in app code (erasure/retention paths exempt); destructive DDL outside a `*contract*` migration |
| `assets/check-isolation-tests.sh` | 28 | a new route file with no nearby test mentioning 404 (`*public*`/`*health*` exempt) |

They are not part of the core gate set: wire them as pre-commit pre-flight steps or CI checks **in repos where the concern exists** (personal data, network IO, a schema, tenants). They are tripwires, not proofs; the discipline references keep the full review duty. The repo smoke test exercises all four so a regression in a guard fails CI here first.

### Never bypass with `--no-verify`

`git commit --no-verify` skips every gate. It is reserved for genuine big-bang changes — initial scaffolds, mass-rename refactors, generated-file updates — never for a failing check. **Justify every bypass in the commit body.** Do not normalise bypassing.

If a check is wrong for the codebase, fix it at the project level — raise or lower a rule's severity in `eslint.config.js`, adjust a coverage gate in `scripts/check-coverage.ts`, refactor a flaky test — and commit the fix. Same discipline as the no-inline-ignore rule: refactor or reconfigure, never suppress.

### Adapt for Husky or another hook manager

If the repo already uses Husky, drop the body of `assets/pre-commit` (from `set -euo pipefail` onwards) into `.husky/pre-commit`. The shebang and the `git config core.hooksPath` step are unnecessary; Husky handles them.

## Dependency CVE scanning (CI)

Gate 2 pins every dependency to a concrete version for supply-chain safety, but a pinned `^1.2.3` can still *be* a known-vulnerable version — pinning stops silent upgrades, it does not scan. The only dependency scanner the toolchain ships (the Snyk IDE extension, see `references/bun-typescript.md`) runs IDE-side, so CI and the pre-commit hook never see its findings — the same drift problem that motivated mirroring SonarLint into ESLint (see *SonarLint findings caught at lint time* above). The fix is a deterministic CVE scan in CI.

**Tool: `bun audit`.** Bun-native, no new dependency, reads the resolved tree from `bun.lock`, and exits non-zero when it lists a vulnerability. `--audit-level=high` filters to high/critical; an unfixable advisory is allow-listed with `--ignore <id>` **at the workflow level, with a reason** — never inline, the same rule as a project-level ESLint severity change.

**It is a CI job, not a ninth gate.** CVE feeds change daily, independent of your diff. Blocking a 10-file commit because a new advisory dropped overnight in an *untouched* dependency fails in the wrong place. So the scan runs in CI on two triggers, each doing a different job:

- **A scheduled daily run** is the real watchdog — it is the only thing that catches a newly-disclosed CVE in a dependency *nobody touched*. A red scheduled run is the signal; wire it to an issue or chat alert if you want (out of scope here).
- **A pull-request run scoped to `package.json` / `bun.lock`** blocks vulnerabilities a PR *deliberately introduces*, while never red-flagging PRs that don't change dependencies.

`.github/workflows/audit.yml`:

```yaml
name: audit
on:
  schedule:
    - cron: '0 6 * * *'        # daily watchdog: new CVEs in untouched deps
  pull_request:
    paths:                      # PR run fires only when deps actually change
      - package.json
      - bun.lock
jobs:
  audit:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: oven-sh/setup-bun@v2
      - run: bun install --frozen-lockfile   # also fails on lockfile drift — a supply-chain check in itself
      - run: bun audit --audit-level=high
      # An advisory with no upstream fix is allow-listed here, project-level, with a reason:
      # - run: bun audit --audit-level=high --ignore GHSA-xxxx-xxxx-xxxx  # no patched release as of YYYY-MM-DD; tracking <link>
```

`--audit-level=high` fails the job only on high/critical advisories; moderate and low are reported but do not block — run `bun audit` locally to see the full list. The scan covers **all** dependencies, not `--prod` only: dev and build tooling are part of the supply-chain attack surface CI exists to watch.

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
- **A skill description is a triggering contract; edits to it rerun the trigger eval.** Any change to a `SKILL.md` frontmatter description runs its eval set before landing (`bash scripts/trigger-eval/run.sh <set> <skill-dir>`; the `suite-routing.json` set with `TRIGGER_EVAL_SUITE` when wording could shift which suite skill wins a query). A description tuned by feel regresses silently; the eval is one command.

## README consistency

The README is the contract with anyone who clones the repo. If it lies, the change is broken even if the tests are green. Audit it twice: once before declaring a task done, and once more before ending the session.

### When to audit

Run the audit whenever the working tree has uncommitted changes on a non-`README.md` file. Skip it only when the changes are clearly internal-only — private helpers, test-only refactors, formatting passes, dependency bumps that do not change usage.

### What to walk

The user-visible surface area is the set of facts the README documents about the project from the outside. For an atelier-shaped repo, that is roughly:

| Surface | What changed in the session that would invalidate the README |
|:---|:---|
| Install / setup steps | Added a system dep (`gitleaks`, `bun`), changed the install command, moved a config file the install copies |
| `package.json` scripts | Added/renamed/removed any of `test`, `lint`, `typecheck`, `coverage`, `mutate:*`, etc.; changed what one of them does |
| CLI flags / subcommands | New flag, renamed flag, changed default, removed flag — both the flag itself and the example invocations in the README |
| Env vars / config files | New `process.env.X` read in `src/composition/env.ts`; new entry in `.env.example`; new key in `bunfig.toml` |
| Top-level layout / architecture diagram | New top-level folder, renamed folder, deleted folder — the README's tree diagram and any prose that names paths |
| Public exports | A function/type/module the README documents as the API surface (not the same as "everything exported from `src/`") |
| Pinned versions | The README mentions "Bun ≥ X" or "Next.js Y" and the actual `package.json` / `bunfig.toml` pin moved |

If the audit finds drift, fix the README in the **same commit** as the code change — drifted READMEs across separate commits are how docs rot.

### Past breakages this rule catches

- CLI flags renamed but `--flow` examples stayed
- Scripts added to `package.json` but not listed
- Folders deleted (or renamed) but the architecture diagram still referenced them
- Coverage and prompt sections missing entirely from a change set that introduced them
- Install one-liner stayed pointing at a deprecated tool while the docs body listed the new one

### Five-check task-done gate

The four inner-loop checks from the top of this file (`bun test`, `bun run lint`, `bun run typecheck`, `bun run coverage`) plus a fifth: `README.md` audited against the surface table above — either updated, or a one-sentence "nothing user-visible changed".

### End-of-session re-audit

A session usually contains several back-to-back tasks. Each one might pass its task-done audit, then the next one drifts the README again. So re-walk the surface table once more before stopping the session — even if every individual task said "nothing user-visible changed", the cumulative diff often does. State the result in one sentence: "README still current" or "README updated for X, Y, Z".

## Editor configuration that keeps formatting stable

Two guardrails prevent Prettier ↔ VS Code TS-formatter drift:

1. `source.fixAll.eslint` on save applies the ESLint-with-Prettier rules **after** whatever formatter handled the file, so the lint rules always have the last word. TS/TSX files format with `vscode.typescript-language-features` (its output is then normalised by the ESLint fix pass); everything else defaults to `dbaeumer.vscode-eslint`. The canonical per-variant `.vscode/settings.json` blocks live in `references/bun-typescript.md` and `references/nextjs-monorepo.md`.
2. The pre-commit hook runs the full four-check loop, catching any drift at commit time.

```json
// .vscode/settings.json (excerpt — full blocks in the variant references)
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

- **Inner-loop checks, always, in order:** `bun test`, `bun run lint`, `bun run typecheck`, `bun run coverage`.
- **Zero warnings, zero inline ignores.** Refactor or change severity at the project level; never suppress per-line.
- **Coverage gates per-tier:** 100% on `domain` + `use-cases`, 80% on `composition` + `infra` + `presenter`, skip `test-helpers` and `main.ts` only. `build-deps.ts` is now in scope (testable via optional config DI).
- **SonarLint parity at lint time** via `eslint-plugin-sonarjs` + type-aware `@typescript-eslint` rules.
- **Pre-commit hook runs the fast gates** (commit size, package.json, gitleaks protect, lint:staged, typecheck); **CI (`assets/ci.yml`) runs the full set** and is the required merge check: strict lint, typecheck, the whole test suite, coverage, and mutation, on a frozen lockfile, plus `bun audit`.
- **Commit identity** (rule 26): contributor identity in commit metadata is normal and never a finding; file contents never name a person, an employer, or a client. Scrubbing a mention from pushed history takes a gated `git filter-repo` rewrite plus a force-push, and the host may keep the old commits cached.
- **Dependency CVE scanning lives in CI, not the gate** (`bun audit --audit-level=high`): a daily scheduled watchdog for new CVEs in untouched deps, plus a PR run scoped to `package.json` / `bun.lock` for deliberately-introduced ones.
- **Mutation testing on staged files** (Stryker, ≥90% break threshold) makes "tests don't actually pin behaviour" findable in CI.
- **Verification discipline:** a control is a hypothesis until a test walks the forbidden path; test the bypass, audit the seams, fix the class not the instance, and proof is a runnable check, not a checklist or a screenshot. Generated code meets the identical bar; provenance is not proof.
- **Commits stay small:** ≤10 files AND ≤300 lines per commit. The hook enforces it.
- **Periodic audits**: once per release, drop the `test-helpers` skip and run coverage; anything below 100% is dead code or untested defensive code.
- **README.md is part of the change set.** Re-read it before declaring any task done.
