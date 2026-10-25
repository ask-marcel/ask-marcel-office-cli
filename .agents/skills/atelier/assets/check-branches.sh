#!/usr/bin/env bash
#
# Hard rule 38's watchdog: no branch outlives its landing, and none lives past a day.
#
# Lists every branch on the remote except the default branch and fails when one
#   - has landed: merging it into the default branch would change nothing, judged
#     by content (git merge-tree), so a rebase or squash merge counts as landed even
#     though its commits are not ancestors of the default branch;
#   - has stayed off the trunk too long: its oldest commit not on the default branch
#     was authored more than MAX_BRANCH_AGE_HOURS ago (default 24, the trunk-based
#     limit of references/workflow.md, Branch lifecycle).
# KEEP_BRANCHES is an extended regex of branch names to leave alone (default
# ^release/: a release branch is cut from the trunk when a release ships and lives on).
# Not a merge gate: a branch goes stale while nobody pushes, so the shipped
# branches.yml runs this on a schedule, and the owner deletes what it lists.
#
# Usage:
#   bash scripts/check-branches.sh             # the remote's branches, fetched first
#   bash scripts/check-branches.sh --selftest  # prove it fails on each case
# Env: DEFAULT_BRANCH (default: the remote's HEAD, else main), MAX_BRANCH_AGE_HOURS,
#      KEEP_BRANCHES, REMOTE (default origin). Needs git 2.38+ (merge-tree --write-tree).
set -euo pipefail

if [ "${1:-}" = "--selftest" ]; then
  tmp=$(mktemp -d)
  trap 'rm -rf "$tmp"' EXIT
  gate=$(cd "$(dirname "$0")" && pwd)/$(basename "$0")
  git init -q --bare "$tmp/origin.git" && git -C "$tmp/origin.git" symbolic-ref HEAD refs/heads/main
  git clone -q "$tmp/origin.git" "$tmp/work" 2>/dev/null && cd "$tmp/work"
  git config user.email t@e.st && git config user.name t && git config push.negotiate false && git checkout -q -b main
  echo base > base.txt && git add -A && git commit -qm "chore: base" && git push -q origin main
  branch() { # $1 name, $2 hours since its commit was authored; one commit off main
    git checkout -q -b "$1" main && echo "$1" > "$(echo "$1" | tr / -).txt" && git add -A
    GIT_AUTHOR_DATE="@$(( $(date +%s) - $2 * 3600 )) +0000" git commit -qm "feat: $1"
    git push -q origin "$1" && git checkout -q main
  }
  expect() { # $1 what, $2 exit wanted, $3 text the output must hold ('' for none), $4 text it must not hold
    local out rc=0
    out=$(env "${@:5}" bash "$gate" 2>&1) || rc=$?
    if [ "$rc" -ne "$2" ] || { [ -n "$3" ] && [[ "$out" != *"$3"* ]]; } || { [ -n "$4" ] && [[ "$out" == *"$4"* ]]; }; then
      echo "selftest FAIL: $1 (exit $rc):" >&2; echo "$out" >&2; exit 1
    fi
  }
  branch fresh 1
  expect "a fresh branch failed" 0 "1 open branch" ""
  branch rebased 0 && branch squashed 0 && git checkout -q squashed && echo more > more.txt \
    && git add -A && git commit -qm "feat: squashed, part two" && git push -q origin squashed && git checkout -q main
  git cherry-pick rebased >/dev/null && git merge -q --squash squashed >/dev/null 2>&1 && git commit -qm "feat: squashed" && git push -q origin main
  branch old 30 && branch release/1.0 100
  expect "a rebase-merged branch passed" 1 "LANDED  rebased" "fresh"
  expect "a squash-merged branch passed" 1 "LANDED  squashed" ""
  expect "a day-old branch passed" 1 "STALE  old" "release/1.0"
  expect "an empty KEEP_BRANCHES still exempted a branch" 1 "STALE  release/1.0" "" KEEP_BRANCHES=
  expect "MAX_BRANCH_AGE_HOURS was ignored" 1 "LANDED  rebased" "STALE  old" MAX_BRANCH_AGE_HOURS=48
  git push -q origin --delete rebased squashed old
  expect "the cleaned remote failed" 0 "1 open branch" "main"
  echo "selftest OK: rejects a rebase-merged and a squash-merged branch left on the remote and a day-old one, keeps release/ unless KEEP_BRANCHES says otherwise, passes a fresh branch"
  exit 0
fi

remote="${REMOTE:-origin}"
max_hours="${MAX_BRANCH_AGE_HOURS:-24}"
keep="${KEEP_BRANCHES-^release/}"

if [ "$(git rev-parse --is-shallow-repository)" = "true" ]; then
  git fetch --quiet --unshallow "$remote" || { echo "check-branches: cannot unshallow $remote" >&2; exit 2; }
fi
git fetch --quiet --prune "$remote" "+refs/heads/*:refs/remotes/$remote/*" \
  || { echo "check-branches: cannot fetch $remote's branches" >&2; exit 2; }
default="${DEFAULT_BRANCH:-}"
if [ -z "$default" ]; then
  default=$(git symbolic-ref --quiet --short "refs/remotes/$remote/HEAD" 2>/dev/null || true)
  default="${default#"$remote"/}"
fi
default="${default:-main}"
trunk="refs/remotes/$remote/$default"
git rev-parse --verify -q "$trunk" >/dev/null || { echo "check-branches: no $remote/$default to compare against" >&2; exit 2; }
trunk_tree=$(git rev-parse "$trunk^{tree}")

now=$(date +%s)
landed=0 stale=0 open=0
while IFS= read -r ref; do
  name="${ref#refs/remotes/"$remote"/}"
  if [ "$name" = "$default" ] || [ "$name" = HEAD ]; then continue; fi
  if [ -n "$keep" ] && [[ "$name" =~ $keep ]]; then continue; fi
  merged=$(git merge-tree --write-tree "$trunk" "$ref" 2>/dev/null | head -1) || merged=""
  if [ "$merged" = "$trunk_tree" ]; then
    printf '  ╳ LANDED  %s  its work is already on %s; delete it (rule 38)\n' "$name" "$default" >&2
    landed=$((landed + 1))
    continue
  fi
  oldest=$(git log --format=%at "$trunk..$ref" | tail -1)
  hours=$(( (now - ${oldest:-$now}) / 3600 ))
  if [ "$hours" -ge "$max_hours" ]; then
    printf '  ╳ STALE  %s  %s h off %s, past the %s h limit (rule 38)\n' "$name" "$hours" "$default" "$max_hours" >&2
    stale=$((stale + 1))
    continue
  fi
  open=$((open + 1))
done < <(git for-each-ref --format='%(refname)' "refs/remotes/$remote/")

if [ $((landed + stale)) -gt 0 ]; then
  {
    echo ""
    [ "$landed" -eq 0 ] || echo "  $landed landed branch(es): delete each (git push $remote --delete <branch>) and turn on the host's automatic head-branch deletion."
    [ "$stale" -eq 0 ] || echo "  $stale branch(es) past ${max_hours} h: land them, split them, or hide the unfinished work behind a flag on $default."
    echo "  Hard rule 38 (references/workflow.md, Branch lifecycle)."
  } >&2
  exit 1
fi
echo "check-branches: $open open branch(es) on $remote, none landed or older than ${max_hours} h"
