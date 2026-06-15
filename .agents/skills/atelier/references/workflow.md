# Workflow

The four-check loop, coverage gates, lint discipline, and the editor/CI rules that keep them enforced. Run through this after every code change; nothing ships until it is clean.

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

`bun run lint:strict` (~25 s) sets the env var `LINT_STRICT=1`; the same `eslint.config.js` reads `process.env['LINT_STRICT']` and conditionally adds a type-aware block (`parserOptions.projectService: true` plus `@typescript-eslint/no-unnecessary-type-assertion` and `@typescript-eslint/prefer-promise-reject-errors`). One config file, two modes — no separate `eslint.strict.config.js` to keep in sync. The pre-commit hook (gate 5) runs the strict version; trust the hook for the inner loop.

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

Two scripts handle this. Both ship in the skill at `assets/`:

| Script | Job |
|---|---|
| `assets/coverage-preload.ts` | Reference template; not used directly when the auto-regenerate flow is wired up |
| `assets/regenerate-coverage-preload.ts` | Walks `src/{infra,composition,presenter}/`, excludes `*.test.ts` / `ports/` / `index.ts`, writes a fresh `scripts/coverage-preload.ts` |

**Two modes:**

```bash
# Write a fresh preload from the current src/ tree.
bun run scripts/regenerate-coverage-preload.ts

# Exit non-zero if the on-disk file is out of sync — for pre-commit / CI.
bun run scripts/regenerate-coverage-preload.ts --check
```

**Optional: wire `--check` into pre-commit as a pre-flight check** before the eight numbered gates. It's <100 ms and an out-of-sync preload silently lies about coverage:

```bash
# In .githooks/pre-commit (or assets/pre-commit), before "[1/8] commit size":
echo "pre-flight: coverage-preload sync" >&2
bun run scripts/regenerate-coverage-preload.ts --check
```

The pre-flight check is **not part of the eight numbered gates** — those stay 1..8 to keep the branding stable. The pre-flight is "you forgot to regenerate after adding a new file"; treat it like a typo check, not a quality gate.

Then the workflow becomes: add a new file under `src/infra/`, run `bun run scripts/regenerate-coverage-preload.ts`, stage both files together, commit. The `--check` invocation catches any forgotten regeneration before it merges. CI can run the same `--check` as a separate job.

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

The fix is a preload file that side-effect-imports every infra, composition, and presenter module, so they appear in the coverage table at 0% if no test exercises them. A ready-to-copy template lives in the skill at `assets/coverage-preload.ts`.

```ts
// scripts/coverage-preload.ts
// This file forces every module that belongs under a coverage gate to appear
// in `bun test --coverage` output, even when no test imports it. Without this,
// an untested adapter is silently absent and the per-file gate passes.
//
// RULE: every new file in src/infra/, src/composition/, or src/presenter/
// must be added here in the same commit. Enforced by code review.

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

**Maintenance rule:** every new file in `src/infra/`, `src/composition/`, or `src/presenter/` must be added to `coverage-preload.ts` in the same commit. Reviewers check this explicitly. A pre-commit lint could enforce it automatically; not done yet, so it is a review obligation.

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

In `eslint.config.js` (one config, two modes — `LINT_STRICT=1` switches on the type-aware block):

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

- **Every commit keeps `main` releasable.** The eight-gate pre-commit hook (below) is the gate that lets you commit to a shared trunk without breaking everyone else — tests pass, types check, coverage and mutation hold, no secrets, before the commit lands.
- **Commits stay small** (gate 1: ≤10 files AND ≤300 lines). Small commits are the unit of continuous integration; they review in minutes, revert cleanly, and bisect precisely. A 300-line ceiling is a trunk-based ceiling.
- **History stays linear and legible** (Conventional Commits, `commit-msg` hook). A trunk read top-to-bottom is the changelog.
- **Incomplete work hides behind a flag, not a branch.** When a feature spans several commits, keep each commit green and the half-built path dark behind a feature flag or simply unreferenced — never park weeks of work on a divergent branch. This is the same instinct as YAGNI and "minimal": ship the smallest safe increment.

Practical loop: pull/rebase often to stay close to the trunk; run the four-check loop after every change; when green, propose the commit and commit it once the user confirms (SKILL.md hard rule 25 — the agent never commits or pushes on its own initiative); then push on their say-so. If a change is too big to land safely in one ≤300-line commit, split it into a sequence of green commits, not a long-lived branch. Releases are cut from the trunk (tag or release branch at the moment of release), never developed on for weeks beforehand.

This is the default for this codebase. It overrides any tooling habit of "branch first by default" — branch only when a short-lived branch genuinely helps (e.g. a PR-review gate your team requires), and merge it the same day.

## Pre-commit hook (eight gates)

The hook is the safety net for the entire workflow. It runs **eight gates** in cost-ascending order — cheap fast-fail gates first, expensive gates last so a slow mutation run only happens when everything else is clean.

This eight-gate hook is the **Bun-script variant's** mechanism. The Next.js monorepo uses `simple-git-hooks` (test + lint + commitlint) instead — see `references/nextjs-monorepo.md`. Never install both: `core.hooksPath` and `simple-git-hooks` overwrite each other.

| # | Gate | Purpose | Typical time |
|:--:|:---|:---|:--:|
| 1 | `scripts/check-commit-size.sh` | ≤10 files AND ≤300 lines | <1s |
| 2 | `scripts/check-package-json.sh` | no `"latest"` or `"*"` version strings | <1s |
| 3 | `gitleaks protect --staged` | secret scan on the staged diff | ~50ms |
| 4 | `bun test` | unit tests pass | seconds |
| 5 | `bun run lint:strict` | type-aware ESLint, 0 errors AND 0 warnings | ~25s |
| 6 | `bun run typecheck` | `tsc --noEmit` clean | seconds |
| 7 | `bun run coverage` | per-tier thresholds pass | seconds |
| 8 | `bun run mutate:staged` | ≥90% mutation score on staged domain/use-case files | 1–3 min per staged file |

A ready-to-copy script lives in the skill at `assets/pre-commit`. The companion scripts (`check-commit-size.sh`, `check-package-json.sh`, `mutate-staged.sh`, `mutate-changed.sh`, `regenerate-coverage-preload.ts`) live alongside it, plus `assets/commit-msg` — a separate git hook documented under *Commit message format* below.

### Install once per clone

```bash
mkdir -p .githooks scripts
cp <skill>/assets/pre-commit .githooks/pre-commit
cp <skill>/assets/commit-msg .githooks/commit-msg
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

`core.hooksPath .githooks` picks up **both** `.githooks/pre-commit` (the eight gates, on the staged diff) and `.githooks/commit-msg` (Conventional Commits, on the message) — one config, two hooks.

Add to `package.json`:

```json
{
  "scripts": {
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

`scripts/check-commit-size.sh` blocks any commit exceeding **10 files OR 300 lines** (insertions + deletions). The thresholds are conservative because they force the discipline; loosening them undermines the rule.

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

### Mutation testing with Stryker (gate 8)

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
- **`bun run mutate:staged`** — files staged for the next commit. Used by gate 8. Skips with exit 0 when no relevant files are staged, so commits to docs, tests, or scripts are unaffected.

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

The eight gates above run on the **staged diff** via `pre-commit`. Commit *messages* are validated by a separate git hook, `commit-msg`, which fires after you write the message — `assets/commit-msg`, installed to `.githooks/commit-msg` and picked up by the same `core.hooksPath`. It is not a ninth gate; it is a different hook on a different input (SKILL.md hard rule 23).

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

### Never bypass with `--no-verify`

`git commit --no-verify` skips every gate. It is reserved for genuine big-bang changes — initial scaffolds, mass-rename refactors, generated-file updates — never for a failing check. **Justify every bypass in the commit body.** Do not normalise bypassing.

If a check is wrong for the codebase, fix it at the project level — raise or lower a rule's severity in `eslint.config.js`, adjust a coverage gate in `scripts/check-coverage.ts`, refactor a flaky test — and commit the fix. Same discipline as the no-inline-ignore rule: refactor or reconfigure, never suppress.

### Adapt for Husky or another hook manager

If the repo already uses Husky, drop the body of `assets/pre-commit` (from `set -euo pipefail` onwards) into `.husky/pre-commit`. The shebang and the `git config core.hooksPath` step are unnecessary; Husky handles them.

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

1. `bun test` — passes
2. `bun run lint` — 0 errors AND 0 warnings (fast, ~2 s cached / ~7 s cold)
3. `bun run typecheck` — clean
4. `bun run coverage` — per-directory gates pass
5. `README.md` — audited against the surface table above; either updated, or one-sentence "nothing user-visible changed"

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
- **Pre-commit hook runs eight gates**, in cost-ascending order: commit size → package.json (no `"latest"` / `"*"`) → gitleaks → tests → lint:strict → typecheck → coverage → mutate:staged.
- **Mutation testing on staged files** (Stryker, ≥90% break threshold) makes "tests don't actually pin behaviour" findable in CI.
- **Commits stay small:** ≤10 files AND ≤300 lines per commit. The hook enforces it.
- **Periodic audits**: once per release, drop the `test-helpers` skip and run coverage; anything below 100% is dead code or untested defensive code.
- **README.md is part of the change set.** Re-read it before declaring any task done.
