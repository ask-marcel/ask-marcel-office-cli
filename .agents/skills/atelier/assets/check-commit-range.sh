#!/usr/bin/env bash
#
# CI mirror of the pre-commit commit-size gate (scripts/check-commit-size.sh).
#
# The hook inspects `git diff --cached`, one prospective commit. CI has no
# staged index, so this walks every non-merge commit a push or pull request
# adds and holds each to the SAME thresholds: <=10 files AND <=300 lines
# (insertions + deletions). The rule is PER COMMIT, not cumulative: a
# 30-commit pull request is fine, one 400-line commit is not, and neither can
# hide inside a squashed range. This is to check-commit-size.sh what
# check-commit-messages.sh is to the commit-msg hook: the half that a
# `--no-verify` bypass cannot skip.
#
# Usage: bash scripts/check-commit-range.sh [base] [head]
#   Defaults: base = origin/$GITHUB_BASE_REF (pull request) or HEAD~1 (push),
#             head = HEAD. Never walks the whole history.
#
# Merge commits are excluded: a merge legitimately touches many files.
# Keep MAX_FILES / MAX_LINES in lockstep with check-commit-size.sh.
#
# Adopted from a consumer repo that had written it independently, found by the
# 2026-08-30 field test (field-test.md). See references/workflow.md
# (Commit size limits).
set -euo pipefail

# --selftest builds a throwaway repo and proves the gate both ways, so the gate
# can be trusted on a machine that is not running the full smoke suite
# (canon 15.10: a gate only ever seen green is a hypothesis).
if [ "${1:-}" = "--selftest" ]; then
  tmp=$(mktemp -d)
  trap "rm -rf '$tmp'" EXIT
  gate=$(cd "$(dirname "$0")" && pwd)/$(basename "$0")
  cd "$tmp"
  git init -q . && git config user.email t@e.st && git config user.name t
  echo one > a.txt && git add -A && git commit -qm "chore: base"
  echo two > b.txt && git add -A && git commit -qm "feat: small change"
  if ! bash "$gate" HEAD~1 HEAD >/dev/null; then
    echo "selftest FAIL: a small commit was rejected" >&2; exit 1
  fi
  i=0; while [ "$i" -lt 12 ]; do printf 'x\n%.0s' $(seq 1 40) > "big$i.txt"; i=$((i + 1)); done
  git add -A && git commit -qm "feat: oversized"
  if bash "$gate" HEAD~1 HEAD >/dev/null 2>&1; then
    echo "selftest FAIL: an oversized commit was accepted" >&2; exit 1
  fi
  # a merge legitimately touches many files and is excluded by design
  git checkout -q -b side HEAD~1 && echo s > s.txt && git add -A && git commit -qm "feat: side"
  git checkout -q - && git merge -q --no-ff side -m "Merge branch 'side'" >/dev/null 2>&1
  if ! MAX_FILES=1000 MAX_LINES=100000 bash "$gate" HEAD~1 HEAD >/dev/null; then
    echo "selftest FAIL: a merge commit was not excluded" >&2; exit 1
  fi
  echo "selftest OK: gate rejects an oversized commit, accepts a small one, excludes merges"
  exit 0
fi

MAX_FILES="${MAX_FILES:-10}"
MAX_LINES="${MAX_LINES:-300}"

base="${1:-}"
head="${2:-HEAD}"

if [ -z "$base" ]; then
  if [ -n "${GITHUB_BASE_REF:-}" ] && git rev-parse --verify -q "origin/$GITHUB_BASE_REF" >/dev/null; then
    base="origin/$GITHUB_BASE_REF"
  elif git rev-parse --verify -q HEAD~1 >/dev/null; then
    base="HEAD~1"
  else
    echo "commit-range: single-commit history, nothing to compare" >&2
    exit 0
  fi
fi

violations=0
for sha in $(git rev-list --no-merges "${base}..${head}"); do
  files=$(git show --pretty="" --name-only --diff-filter=ACMR "$sha" | grep -c '^' || true)
  lines=$(git show --pretty="" --numstat "$sha" | awk '{ sum += $1 + $2 } END { print sum + 0 }')
  if [ "${files:-0}" -gt "$MAX_FILES" ] || [ "${lines:-0}" -gt "$MAX_LINES" ]; then
    printf '  ╳ COMMIT TOO BIG  %s  %s files / %s lines  (max %s / %s)\n' \
      "$(git rev-parse --short "$sha")" "$files" "$lines" "$MAX_FILES" "$MAX_LINES" >&2
    printf '      %s\n' "$(git show -s --format=%s "$sha")" >&2
    violations=$((violations + 1))
  fi
done

if [ "$violations" -gt 0 ]; then
  {
    echo ""
    echo "  $violations commit(s) exceed the atelier size limit (pre-commit gate 1, canon 8.1)."
    echo "  Split each into <=300-line slices; an unreviewable commit is an unreviewed commit."
  } >&2
  exit 1
fi

echo "commit-range: every commit in ${base}..${head} is within ${MAX_FILES} files / ${MAX_LINES} lines"
