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

# Stryker's `--mutate` is a single overrideable flag (`commander`-style):
# repeating `--mutate path1 --mutate path2` keeps only the LAST value.
# Join the file list into one comma-separated argument, the documented
# form for narrowing the mutation set to multiple files.
mutate_arg=$(echo "$files" | paste -sd, -)

bunx stryker run --mutate "$mutate_arg"
