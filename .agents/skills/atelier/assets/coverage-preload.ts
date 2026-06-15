/*
 * Coverage preload.
 *
 * `bun test --coverage` only reports rows for files the runner imports.
 * Untested infra, composition, and presenter files are silently absent
 * from the table, which makes the per-file gate trivially pass. This
 * preload side-effect-imports every such file so they appear at 0% (or
 * better) and the gate can fail loudly.
 *
 * Wired ONLY at coverage time, NOT in bunfig.toml. `scripts/check-coverage.ts`
 * spawns `bun test --coverage --preload ./scripts/coverage-preload.ts`. We do
 * NOT put `preload = [...]` under `[test]` in bunfig.toml because this file
 * pulls in heavy third-party SDKs (HTTP clients, cloud SDKs, AI clients,
 * loggers, etc. — whatever the infra adapters wrap) that would slow every
 * plain `bun test` by 1–2s.
 *
 * MAINTENANCE RULE: every new file under
 *   - src/infra/
 *   - src/composition/
 *   - src/presenter/
 * must be side-effect-imported here in the SAME commit that adds the file.
 * Reviewers check this explicitly. A pre-commit lint could enforce it; not
 * done yet, so it is a review obligation.
 *
 * See skills/atelier/references/workflow.md for the full rationale.
 */

// --- src/infra/ ---
import '../src/infra/logger.ts';
// TODO: add every adapter here as you create it, e.g.:
// import '../src/infra/<your-adapter-1>.ts';
// import '../src/infra/<your-adapter-2>.ts';
// import '../src/infra/<your-adapter-3>.ts';

// --- src/composition/ ---
import '../src/composition/env.ts';
import '../src/composition/build-deps.ts';
// NOTE: build-deps.ts USED to be skipped here. It is now testable end-to-end
// via the optional BuildDepsConfig argument (token-store path + logger as
// optional config; sensible defaults preserve production behaviour). See
// references/architecture.md (Composition root testability).

// --- src/presenter/ ---
import '../src/presenter/cli.ts';
