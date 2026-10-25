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
#   Defaults: base = origin/$GITHUB_BASE_REF (pull request), else
#             $GITHUB_EVENT_BEFORE (a push: every commit it added, exported by
#             the shipped workflow from github.event.before), else HEAD~1;
#             head = HEAD. Never walks the whole history.
#
# A merge commit in the range fails (hard rule 38: main stays linear, a branch
# lands by rebase or fast-forward). The one merge stepped over is GitHub's own:
# a pull_request checkout sits on a synthetic merge of the PR into its base
# ("Merge <sha> into <sha>"), and the PR's commits end at its second parent.
# Keep MAX_FILES / MAX_LINES in lockstep with check-commit-size.sh.
#
# Adopted from a consumer repo that had written it independently, found by the
# 2026-08-30 field test. See references/workflow.md
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
  # rule 38: a merge commit at the tip of the range fails, for its own reason
  git checkout -q -b side HEAD~1 && echo s > s.txt && git add -A && git commit -qm "feat: side"
  git checkout -q - && git merge -q --no-ff side -m "Merge branch 'side'" >/dev/null 2>&1
  out=$(bash "$gate" HEAD~1 2>&1) && { echo "selftest FAIL: a merge commit was accepted" >&2; exit 1; }
  case "$out" in *"MERGE COMMIT"*"rule 38"*) ;; *) echo "selftest FAIL: a merge commit was rejected, but not for rule 38:" >&2; echo "$out" >&2; exit 1 ;; esac
  # GitHub's synthetic pull-request merge at HEAD is stepped over, the PR's own commits are not
  trunk=$(git rev-parse HEAD)
  git checkout -q -b pr "$trunk" && echo p > p.txt && git add -A && git commit -qm "feat: pr change"
  git checkout -q --detach "$trunk" && git merge -q --no-ff pr -m "Merge $(git rev-parse pr) into $trunk" >/dev/null 2>&1
  if ! bash "$gate" "$trunk" >/dev/null 2>&1; then
    echo "selftest FAIL: GitHub's synthetic pull-request merge was counted as a merge commit" >&2; exit 1
  fi
  git checkout -q -b pr2 "$trunk" && echo q > q.txt && git add -A && git commit -qm "feat: pr2 change"
  git checkout -q -b pr2side "$trunk" && echo r > r.txt && git add -A && git commit -qm "feat: pr2 side"
  git checkout -q pr2 && git merge -q --no-ff pr2side -m "Merge branch 'pr2side' into pr2" >/dev/null 2>&1
  git checkout -q --detach "$trunk" && git merge -q --no-ff pr2 -m "Merge $(git rev-parse pr2) into $trunk" >/dev/null 2>&1
  out=$(bash "$gate" "$trunk" 2>&1) && { echo "selftest FAIL: a merge inside a pull request was accepted" >&2; exit 1; }
  case "$out" in *"MERGE COMMIT"*"rule 38"*) ;; *) echo "selftest FAIL: a merge inside a pull request was rejected, but not for rule 38:" >&2; echo "$out" >&2; exit 1 ;; esac
  echo "selftest OK: gate rejects an oversized commit and a merge commit (rule 38), accepts a small one, steps over GitHub's synthetic pull-request merge and still sees a merge inside the pull request"
  exit 0
fi

MAX_FILES="${MAX_FILES:-10}"
MAX_LINES="${MAX_LINES:-300}"

base="${1:-}"
head="${2:-HEAD}"
# GitHub's synthetic merge of a pull request into its base: judge the PR's own commits instead
if [ -z "${2:-}" ] && git rev-parse -q --verify 'HEAD^2' >/dev/null \
  && git log -1 --format=%s HEAD | grep -qE '^Merge [0-9a-f]{40} into [0-9a-f]{40}$'; then
  head="HEAD^2"
fi

zero_sha=0000000000000000000000000000000000000000
if [ -z "$base" ]; then
  if [ -n "${GITHUB_BASE_REF:-}" ] && git rev-parse --verify -q "origin/$GITHUB_BASE_REF" >/dev/null; then
    base="origin/$GITHUB_BASE_REF"
  elif [ -n "${GITHUB_EVENT_BEFORE:-}" ] && [ "$GITHUB_EVENT_BEFORE" != "$zero_sha" ] \
    && git rev-parse --verify -q "${GITHUB_EVENT_BEFORE}^{commit}" >/dev/null; then
    base="$GITHUB_EVENT_BEFORE"   # a push: walk every commit it added, not only the tip
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

merges=0
for sha in $(git rev-list --merges "${base}..${head}"); do
  printf '  ╳ MERGE COMMIT  %s  %s\n' "$(git rev-parse --short "$sha")" "$(git show -s --format=%s "$sha")" >&2
  merges=$((merges + 1))
done

if [ "$violations" -gt 0 ]; then
  {
    echo ""
    echo "  $violations commit(s) exceed the atelier size limit (pre-commit gate 1, canon 8.1)."
    echo "  Split each into <=300-line slices; an unreviewable commit is an unreviewed commit."
  } >&2
fi
if [ "$merges" -gt 0 ]; then
  {
    echo ""
    echo "  $merges merge commit(s) in the range: hard rule 38 keeps main linear."
    echo "  Rebase the branch onto main (git rebase origin/main) and land it by rebase or fast-forward."
  } >&2
fi
[ "$violations" -eq 0 ] && [ "$merges" -eq 0 ] || exit 1

echo "commit-range: every commit in ${base}..${head} is within ${MAX_FILES} files / ${MAX_LINES} lines, and none is a merge (rule 38)"
