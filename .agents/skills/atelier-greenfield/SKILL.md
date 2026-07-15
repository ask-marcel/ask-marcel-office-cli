---
name: atelier-greenfield
description: Stand up a NEW Bun/TypeScript, Next.js, or Java (Quarkus) repo to the atelier standard from zero — scaffold the Clean Architecture (or Atomic Design) layout, copy the gate assets, wire the git hooks, write the build scripts, lay a minimal green walking skeleton, and prove every gate passes before the first commit. Use when starting a fresh repo or a new monorepo package from scratch — say "scaffold a new Bun repo", "bootstrap a new project to the standard", "set up a new atelier repo / Next.js package", or "scaffold a new Java/Quarkus service". Greenfield only. NEVER use it for work inside an existing codebase (adding features, setting up lint or tooling, config changes) — those belong to the main atelier skill, or to atelier-review-me adopt mode for bringing a brownfield repo up to the standard.
---

# Greenfield

Bring a repo into existence already conforming to the atelier standard, green from the first commit. This is the companion skill for the one moment the main standard assumes has already happened: repo birth. The raw material — the per-variant bootstrap checklists and the copyable gate assets — already exists; this skill sequences and executes it so a fresh repo passes all eight gates before a line of feature code is written.

Like atelier-grill-me, it is a focused, on-demand counterpart to the always-on atelier standard: atelier-grill-me owns the pre-decision moment, atelier-greenfield owns the repo-birth moment.

Interaction: terse, direct prose with no filler, praise, or recap; never use em dashes; answer first, reasoning only if it changes the decision; challenge on substance, not question spam; ask only when the answer changes what you produce and you cannot infer it from the repo, the user's files, or what they said, then AskUserQuestion (or the client's structured-options equivalent) with 2-4 concrete options led by your recommendation; one question round max, then proceed on assumptions named inline; confirm once before an irreversible action (commit, push, publish); propose next steps at wrap-up.

## When to use

- Starting a brand-new Bun/TypeScript script repo, or a new package in a Next.js monorepo.
- The user says "scaffold / bootstrap / set up a new repo (or package) to the standard".
- An empty or near-empty directory that should become a conforming repo.

Greenfield only. For a repo that already has code, the main atelier skill applies — do not run this over a populated tree; it scaffolds and would collide. Match intensity to the ask: a throwaway spike does not need the full eight-gate machinery — say so and offer the minimal subset.

## How to run

1. **Confirm variant and target directory — one question, with a recommendation.** Read the shape from the directory: a `packages/*` workspace with `next.config.ts` nearby → Next.js package; a `pom.xml`/`build.gradle` context or an explicit Java ask → Java (Quarkus); otherwise → Bun-script. State your read and confirm before touching anything.
2. **Locate the atelier assets.** The gate scripts and helpers live in the installed `atelier` skill's `assets/` directory (typically `~/.claude/skills/atelier/assets/`). Resolve that path and report it. If it is missing, stop and tell the user to install the `atelier` skill first — atelier-greenfield copies from it, it does not vendor its own copies.
3. **Follow the matching bootstrap checklist verbatim.** Read and execute the numbered checklist for the variant — do not improvise the config:
   - Bun-script → the installed `atelier` skill's `references/bun-typescript.md` (§ Bootstrap checklist): layout, `tsconfig`, flat ESLint config, `Result` + Logger port/adapter/fake, asset copy, coverage gate, Stryker, the two git hooks, `bunfig.toml`.
   - Next.js → the installed `atelier` skill's `references/nextjs-monorepo.md` (§ Bootstrap checklist): package layout, `tsconfig`/`eslint`/`postcss`/`next.config`, `simple-git-hooks`, `globals.css`, the `src/components/{atoms,molecules,organisms}` + `src/page` + `src/lib` + `src/config` tree.
   - Java → the installed `atelier` skill's `references/java-quarkus.md` (§ Bootstrap checklist): Quarkus scaffold with the Maven wrapper, pinned pom with Spotless + JaCoCo tiers + PIT + enforcer, the `domain`/`usecases`/`infra`/`api`/`composition` packages, sealed `Result`, Flyway, and the shipped hook assets (`assets/pre-commit-java`, `assets/check-pom.sh`, the shared `assets/check-commit-size.sh` and `assets/commit-msg`), copied not hand-written.
   Never install both hook mechanisms — eight-gate `.githooks` for Bun-script, `simple-git-hooks` for Next.js; the Java variant reuses `.githooks` with its own gate chain.
4. **Pin versions properly.** Add every dependency with `bun add` / `bun add -d` so it resolves to a concrete `^X.Y.Z`; never hand-write `"latest"` or `"*"` (rule 19 — gate 2 would reject it). Java: exact versions in the pom, no ranges, no SNAPSHOT deps (the enforcer blocks them).
5. **Choose the commit identity deliberately.** Before the first commit, decide what identity every commit will carry (rule 26): a neutral repo-local handle (for example `atelier <atelier@users.noreply.github.com>`) if the repo should not be tied to a person, or your own name and email if you want attribution on your work. Set it explicitly with `git config --local` so it is a conscious choice, not whatever the global identity happens to be. Repo birth is the only moment this is free; changing it later means a `git filter-repo` history rewrite and a force-push.
6. **Seed session memory and the standard pointer.** Create `.claude/LESSONS.md` with just its header so the cross-session journal works from commit one, and write a minimal `CLAUDE.md` so the standard rides in deterministic repo context on every future session instead of depending on skill triggering alone:

   ```markdown
   # CLAUDE.md

   This repo follows the atelier coding standard. Consult the `atelier` skill for every
   code task here; its hard rules 1-34 bind (TDD with hand-written fakes, `Result` at IO
   boundaries, branded types at trust boundaries, and the production disciplines: privacy,
   isolation, reliability, observability). Run the `atelier-review-me` skill before landing
   changes. Journals: `.claude/LESSONS.md` (append-only memory), `.claude/PLAN.md` (current plan).
   ```

   Offer `/init` afterwards to extend it with codebase-specific documentation; the pointer block above stays at the top.
7. **Lay a minimal walking skeleton.** The thinnest end-to-end slice that touches every layer (see the installed `atelier` skill's `references/architecture.md` § The walking skeleton): for a Bun-script repo, one use-case returning `Result.ok` through its primary port, its branded input, and its confirmed test (rule 24 — propose the test, get the yes, then write it); for Next.js, one `src/lib` pure function with a test wired into a page shell that renders a single atom; for Java, one application service returning `Ok` through its port with its JUnit test, plus one resource with its REST Assured test including the 401 refusal. This is not speculative code — it is what makes coverage and mutation pass for real and demonstrates the TDD loop in place. Keep it to the absolute minimum, and offer to skip it for a bare scaffold.
8. **Prove green.** Run the inner loop and confirm each is clean — `bun test`, `bun run lint`, `bun run typecheck`, `bun run coverage`, and for the Bun-script variant `bun run mutate`; for Java, `./mvnw spotless:check verify` plus PIT on the skeleton. Confirm the `commit-msg` hook rejects a junk message. A bootstrap that does not end green has not finished.
9. **Stop before the first commit.** Stage the tree, propose the Conventional-Commits message (`chore: scaffold repo` or similar), and wait for the user's explicit yes (rule 25). Never auto-commit.

## Output

A directory that passes every gate, with a one-screen summary: variant chosen, assets copied (and from where), hooks wired, commit identity chosen, scripts added, the walking skeleton's single scenario, the inner-loop results, and the proposed first-commit message awaiting confirmation.

Scaffold the *simplest* conforming repo that runs green — the walking skeleton is one slice, not a feature. Everything past it is built later, test-first, under the main atelier standard. Out of scope here: brownfield adoption of an existing repo, CI/CD, deployment, Docker, and choosing a web framework.

This skill is the standard's paved road (the golden path): a repo born from it starts already passing every gate, instead of migrating toward compliance later. Ship real artifacts, not instructions; the fastest way to start is also the compliant one. When the new repo will hold personal data, serve multiple tenants, or call an AI model, name the production disciplines that will bind (atelier rules 27-34) in the wrap-up summary so the first features are designed with them, not retrofitted.
