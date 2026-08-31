#!/usr/bin/env bash
#
# Run Stryker mutation testing on files differing from `origin/main` plus any
# uncommitted edits and untracked files. Used during iteration to catch
# surviving mutants before staging.
#
# Override the base ref with the BASE env var; skip the ref refresh with
# MUTATE_NO_FETCH=1 (offline, or when BASE is deliberately stale):
#
#   BASE=HEAD~3 bun run mutate:changed
#
# A repo with no `origin/main` yet (greenfield, before the first push) must set
# BASE to a local ref; an unknown base fails loudly rather than passing empty.
#
# See skills/atelier/references/workflow.md (Mutation testing).

set -euo pipefail

BASE="${BASE:-origin/main}"

# `origin/*` is a LOCAL cache of the remote, moved only by a fetch. Against a
# stale ref, `$BASE...HEAD` still holds commits pushed long ago: the mutation
# set widens by files nobody touched, and the same list reads as "unpushed
# work". Refresh it, but never fail the run on a network error.
if [ -z "${MUTATE_NO_FETCH:-}" ] && [ "${BASE#origin/}" != "$BASE" ]; then
  git fetch --quiet origin "${BASE#origin/}" || true
fi

# An unknown BASE makes every diff below fail, and the `|| true` on the
# pipeline turns that into "no files changed" plus exit 0: a green run that
# mutated nothing. Fail loudly instead.
if ! git rev-parse --verify --quiet "$BASE^{commit}" >/dev/null; then
  echo "mutate:changed: base ref '$BASE' does not exist (fetch it, or set BASE=)" >&2
  exit 1
fi

echo "mutate:changed: base ${BASE} $(git rev-parse --short "$BASE")" \
     "($(git log -1 --format=%cr "$BASE")), HEAD +$(git rev-list --count "$BASE"..HEAD)"

# Files that differ from BASE, plus uncommitted and staged edits, plus
# untracked files. Untracked matters: a brand-new source file appears in NO
# diff, so without it a new domain or use-case file is never mutated and the run
# still exits 0.
changed=$( {
  git diff --name-only --diff-filter=ACMR "$BASE"...HEAD
  git diff --name-only --diff-filter=ACMR HEAD
  git diff --cached --name-only --diff-filter=ACMR
  git ls-files --others --exclude-standard
} | sort -u)

# A changed test file pulls in the source it covers. Test files carry no
# mutants, so a commit touching ONLY tests used to present zero files: the gate
# printed "no files in mutation scope changed" and passed having run nothing.
# That hides the one change mutation testing exists to catch, a test WEAKENED
# against production code nobody edited, and it hid a real 89.14 behind a green
# run on 2026-08-30. Partial by construction: a shared test file with no sibling
# source (commands.test.ts covers ~180 command modules) still maps to nothing.
covered=$(echo "$changed" | grep -E '\.test\.ts$' | sed -E 's/\.test\.ts$/.ts/' || true)

# `-f` filters the mapping's misses (a test whose sibling source does not
# exist); the diffs themselves are already ACMR, so nothing deleted reaches it.
files=$( { echo "$changed"; echo "$covered"; } | sort -u \
  | grep -E '^src/(domain|use-cases)/' \
  | grep -E '\.ts$' \
  | grep -vE '\.test\.ts$' \
  | grep -vE '/ports/' \
  | while IFS= read -r f; do [ -f "$f" ] && echo "$f"; done \
  || true)

if [ -z "$files" ]; then
  echo "mutate:changed: no files in mutation scope changed since ${BASE}"
  exit 0
fi

count=$(echo "$files" | wc -l | tr -d ' ')
echo "mutate:changed: testing ${count} file(s)"

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
