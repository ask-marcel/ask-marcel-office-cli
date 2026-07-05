#!/usr/bin/env bash
#
# Run Stryker mutation testing on files differing from `origin/main` plus
# any uncommitted edits. Used during iteration to catch surviving mutants
# before staging.
#
# Override the base ref with the BASE env var:
#
#   BASE=HEAD~3 bun run mutate:changed
#
# See skills/atelier/references/workflow.md (Mutation testing).

set -euo pipefail

BASE="${BASE:-origin/main}"

# Files that differ from BASE plus uncommitted/staged edits, intersected
# with the mutation scope.
files=$( {
  git diff --name-only --diff-filter=ACMR "$BASE"...HEAD
  git diff --name-only --diff-filter=ACMR HEAD
  git diff --cached --name-only --diff-filter=ACMR
} | sort -u \
  | grep -E '^src/(domain|use-cases)/' \
  | grep -E '\.ts$' \
  | grep -vE '\.test\.ts$' \
  | grep -vE '/ports/' \
  || true)

if [ -z "$files" ]; then
  echo "mutate:changed: no files in mutation scope changed since ${BASE}"
  exit 0
fi

count=$(echo "$files" | wc -l | tr -d ' ')
echo "mutate:changed: testing ${count} file(s) (base: ${BASE})"

# Stryker's --mutate takes ONE comma-separated value; repeated flags
# overwrite each other (the CLI keeps only the last one), so join the list.
mutate_arg=$(echo "$files" | paste -sd, -)

# Run fresh. Stryker's incremental cache keys on source-file hashes, so a
# test-only change (a stronger assertion, same source) does not invalidate it
# and the score reports stale. Clear it so this pre-stage check is trustworthy.
rm -f reports/stryker-incremental.json

bunx stryker run --mutate "$mutate_arg"
