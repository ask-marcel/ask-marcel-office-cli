#!/usr/bin/env bash
#
# Run Stryker mutation testing on STAGED files in the mutation scope
# (src/domain/** and src/use-cases/**, excluding tests and ports).
#
# An optional local pre-check (`bun run mutate:staged`); the enforcing home is
# CI (`mutate:changed` on every pull request and push; the full `mutate` sweep
# is the daily assets/mutation.yml). Never wired into the hook. Skips with exit 0 when no relevant
# files are staged, so commits that touch only docs, tests, or scripts are
# unaffected.
#
# See skills/atelier/references/workflow.md (Mutation testing).

set -euo pipefail

files=$(git diff --cached --name-only --diff-filter=ACMR \
  | grep -E '^src/(domain|use-cases)/' \
  | grep -E '\.ts$' \
  | grep -vE '\.test\.ts$' \
  | grep -vE '/ports/' \
  || true)

if [ -z "$files" ]; then
  echo "mutate:staged: no staged files in mutation scope, skipping"
  exit 0
fi

count=$(echo "$files" | wc -l | tr -d ' ')
echo "mutate:staged: testing ${count} file(s)"

# Stryker's --mutate takes ONE comma-separated value; repeated flags
# overwrite each other (the CLI keeps only the last one), so join the list.
mutate_arg=$(echo "$files" | paste -sd, -)

# Run the GATE fresh. Stryker's incremental cache (incremental:true in the
# config, kept for fast dev `bun run mutate`) keys on source-file hashes, so a
# test-only change, strengthening an assertion without touching the source,
# does not invalidate it and the score reports stale. A commit gate must judge
# the current tree, so --force ignores cached statuses. Prefer it to deleting
# reports/stryker-incremental.json: that path is owned by stryker.conf.json
# via `incrementalFile`, so a config change silently turns the delete into a
# no-op and the stale-score trap returns.
bunx stryker run --force --mutate "$mutate_arg"

# Stryker's break judges the whole run's score, and under --mutate its report keeps earlier results
# for other files in that score: a file at 55.56 passed in a run that totalled 90.00 (2026-10-23).
# The per-file gate holds each file this run mutated to the same break.
set --
while IFS= read -r f; do set -- "$@" "$f"; done <<<"$files"
bun run scripts/check-mutation-score.ts "$@"
