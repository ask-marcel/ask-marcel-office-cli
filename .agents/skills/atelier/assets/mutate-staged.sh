#!/usr/bin/env bash
#
# Run Stryker mutation testing on STAGED files in the mutation scope
# (src/domain/** and src/use-cases/**, excluding tests and ports).
#
# Used by the pre-commit hook (gate 8). Skips with exit 0 when no relevant
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
# test-only change — strengthening an assertion without touching the source —
# does not invalidate it and the score reports stale. A commit gate must judge
# the current tree, so clear the cache before the staged run.
rm -f reports/stryker-incremental.json

bunx stryker run --mutate "$mutate_arg"
