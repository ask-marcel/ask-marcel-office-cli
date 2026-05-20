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

# Stryker's `--mutate` flag takes a single comma-separated pattern list;
# repeating `--mutate` keeps only the LAST value. Join the file list into
# one comma-separated argument so all staged files are mutated together.
mutate_arg=$(echo "$files" | paste -sd, -)

bunx stryker run --mutate "$mutate_arg"
