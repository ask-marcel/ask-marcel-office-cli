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

staged=$(git diff --cached --name-only --diff-filter=ACMR)

# A staged test file pulls in the source it covers. Test files carry no mutants,
# so staging ONLY tests used to skip the gate entirely — exactly when it matters
# most, since a WEAKENED test lowers the score of source nobody edited. Partial
# by construction: a shared test file with no sibling source maps to nothing.
covered=$(echo "$staged" | grep -E '\.test\.ts$' | sed -E 's/\.test\.ts$/.ts/' || true)

# `-f` drops the mapping's misses (a test whose sibling source does not exist).
files=$( { echo "$staged"; echo "$covered"; } | sort -u \
  | grep -E '^src/(domain|use-cases)/' \
  | grep -E '\.ts$' \
  | grep -vE '\.test\.ts$' \
  | grep -vE '/ports/' \
  | while IFS= read -r f; do [ -f "$f" ] && echo "$f"; done \
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

# Stryker's `--incremental` is a VALUELESS flag (`.option('--incremental', ...)`
# in stryker-cli.js), so there is no `--incremental false` and no
# `--no-incremental` to override `incremental: true` in stryker.conf.json:
# passing a value makes Stryker read it as a config-file path and abort with
# `Invalid config file "false"`. Point the cache at a throwaway file instead. A
# file that does not exist has nothing to replay, so the run tests every mutant
# AND the reported score covers only the files this run mutated.
#
# `--force` was not enough: it re-tests every mutant but still READS the shared
# cache, and the score is computed over that ENTIRE report, folding in results
# for files the run never touched. That inflated a real 89.14 to 95.77 on
# 2026-08-30 and is why a local run and CI disagreed. The configured
# `incrementalFile` is left untouched, so a full `bun run mutate` keeps its
# cache and its speed.
incremental_dir=$(mktemp -d "${TMPDIR:-/tmp}/stryker-inc.XXXXXX")
trap 'rm -rf "$incremental_dir"' EXIT

bunx stryker run --incrementalFile "$incremental_dir/incremental.json" --mutate "$mutate_arg"
