---
name: atelier-review-me
description: Review a diff against the atelier standard before it lands — a rule-aware conformance audit that maps each changed file to the hard rules that bind it, cites the exact rule number — or the red flag — a change violates, applies the security false-positive filter, and defers generic correctness bugs to /code-review and mechanical cleanups to /simplify. Use to pre-land-review staged changes, a feature branch, or a PR, to check a diff for rule violations, or when the user says "review me" / "review my changes against the standard". Also runs an adopt mode for brownfield — scan a whole existing repo and emit a staged plan to bring it up to the standard, e.g. "adopt the standard into this repo", "migrate this repo to atelier", "bring this repo up to standard".
---

# Review me

Audit a change against the atelier standard before it lands. This is the conformance lens the always-on standard and the generic review tools do not give you on their own: a whole-diff pass that checks every changed file against the hard rules it is bound by, in domain language, citing rule numbers — so a violation is caught at review cost, not in production or three rounds into a reviewer's comment thread.

The third on-demand companion to the always-on atelier standard: atelier-grill-me owns the pre-decision moment, atelier-greenfield owns repo-birth, atelier-review-me owns the pre-land moment.

## When to use

- The user asks to "review me", to review a diff / branch / PR against the standard, or to check changes for rule violations before committing.
- A change is staged or a feature branch is ready to land and you want a conformance checkpoint.
- The user wants to adopt the standard into an **existing** repo, migrate a brownfield codebase, or get a staged plan to bring a repo up to the standard — this triggers adopt mode (below).

Match intensity to stakes — a one-line typo fix does not need the full rule sweep; say so and skip it. And atelier-review-me does not replace the built-ins, it complements them (as the security reference complements `/security-review`): defer generic correctness bugs to `/code-review` and mechanical reuse/simplification/altitude cleanups to `/simplify`. atelier-review-me owns rule conformance.

## How to run

1. **Resolve the diff scope — one question, with a recommendation.** Staged (`git diff --cached`), the working tree, the branch vs `origin/main` (`git diff origin/main...`), a PR, or the whole repo (adopt mode — see below). Default recommendation: the current branch vs `origin/main`. Confirm, then read the actual diff before judging.
2. **Map each changed file to its rule subset.** The rules are layer-specific — audit only what applies:
   - `src/domain/**`, `src/use-cases/**` → rules 1-3, 6, 10, 12, 14, 16-18: branded types at trust boundaries, primary-port SUT, `Result` returns, no `try/catch`, no custom error classes.
   - `src/infra/**` → 13, 17, 20: the test seam (a `createXFromApi`/custom-fetch/sync-builder seam, never `mock`), `try/catch` quarantined here, `Bun.file` not `node:fs`, `Result` translation via `formatError`.
   - `src/components/**`, `src/page/**`, `app/**` → 21-22: design-system purity and the styling seal — no hooks, no `src/lib`/`next/*` imports in components, no Tailwind outside `src/components/**`, typed variants not `className`.
   - `*.test.ts`, `src/test-helpers/**` → 24, 13, 14: test integrity (was a test created/edited/weakened without sign-off?), no `mock` from `bun:test`, primary-port SUT, domain-language scenario names.
   - `package.json` → 19 (no `"latest"`/`"*"`). Any commit → 23 (Conventional Commits) and small (≤10 files / ≤300 lines).
3. **Check the universal hard rules on every file** — no `class` / `function` declaration / `interface` / `console.*` (1-4), explicit return types on exports (6), single-arrow not curried (18), zero inline ignores of any tool (15).
4. **Run the security source-to-sink lens.** Does any untrusted source reach a sink (SQL, shell, filesystem, HTTP, HTML, redirect) without crossing a branded-type checkpoint? Apply the strict false-positive filter — only concrete, exploitable findings with a clear attack path; skip DoS, defence-in-depth hardening, and theoretical concerns.
5. **Note, do not re-run, the mechanical gates.** Tests / lint / typecheck / coverage / mutation are enforced by pre-commit gates 4-8 — remind the user to run them rather than checking by eye. atelier-review-me is for the judgment rules the gates cannot catch.

## Adopt mode (brownfield)

When the target is an **existing non-conforming repo** rather than a diff, atelier-review-me switches to adopt mode: it scans the whole tree, then sequences the migration so the repo reaches the standard in green increments instead of one unreviewable rewrite. The conformance scan is the same layer→rule mapping as above, widened from the diff to all of `src/**`; the new work is the *order*. Most of adoption is deciding what NOT to migrate yet — YAGNI applies to migrations too.

1. **Assess.** Scan the whole tree and report which rules are violated, where, and how pervasively. Rank by leverage — the toolchain and the test seam before cosmetic rules.
2. **Install the gates without tripping them on the legacy tree.** Add the ESLint flat config, `tsconfig`, and the hooks, but scope enforcement to changed files at first (lint the diff, not the thousand pre-existing warnings) so every commit isn't blocked. The only sanctioned bypass is the initial mechanical scaffold / mass-rename commit — `--no-verify` with a justification in the commit body (`references/workflow.md` § Never bypass) — never on a failing check, never as a habit.
3. **Characterise before you change.** For each slice you will touch, propose characterisation tests that pin current behaviour (rule 24 — propose, confirm, then write) so the refactor is provably behaviour-preserving. No characterisation test, no refactor.
4. **Migrate slice by slice, each a green ≤300-line commit.** One vertical slice at a time: `class` → module, raw primitive → branded type at the trust boundary, IO → `Result` (per `references/result-type.md` § Migration checklist), `console.*` → the Logger port. The commit-size gate is the unit of work, not an obstacle.
5. **Trunk, not branches.** Land each slice on `main`, half-done work dark behind a flag (`references/workflow.md` § Trunk-based) — never a long-lived migration branch that rots.
6. **Flip the gates to blocking** once the tree conforms: full `lint:strict`, coverage tiers, mutation. Adoption is done when a fresh clone passes all eight gates with no scoping and no bypass.

The output here is a **staged adoption plan** — the ordered slices with the first one ready to start — not a verdict on one diff. It is the brownfield counterpart to what the atelier-greenfield skill does for a new repo: atelier-greenfield births a conforming repo, adopt mode walks an existing one to conformance.

## Output

A rule-cited verdict: each finding names the file, the exact rule number (or the red flag) it breaks, why, and the fix — grouped by severity, in domain language, the single most important fix first. End with a one-line verdict: conformant, or N violations across M files.

Report only — never edit the tree. Offer to apply the fixes on request, hand mechanical cleanups to `/simplify`, and pass correctness bugs to `/code-review`. Review toward the *simplest* conforming change: a finding that demands more code than the rule requires is itself a smell.

In a repo that keeps an `.claude/LESSONS.md` journal, a violation that keeps recurring is a candidate `[mistake]` entry — propose it on approval so the next session inherits the correction.
