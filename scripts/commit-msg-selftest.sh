#!/usr/bin/env bash
#
# Proves the commit-msg hook (.githooks/commit-msg) can fail: each case writes a
# message file the way git does and runs the hook on it. A hook that only ever
# passed would look the same as one that was never installed, so every
# rejection the grammar promises has a case here.
#
# Why the cases exist:
#   - the repo's own headers (a scope, a breaking `!`) and git's generated ones
#     (Merge, Revert, fixup!) must pass, or the hook trains --no-verify;
#   - the comment lines git appends to the template are not the header;
#   - no type, an unlisted type, a capitalised type, a trailing period, a header
#     past 100 characters and an empty message must each be rejected.
#
# Run: bash scripts/commit-msg-selftest.sh

set -euo pipefail

hook="$(cd "$(dirname "$0")/.." && pwd)/.githooks/commit-msg"
failures=0
work=$(mktemp -d "${TMPDIR:-/tmp}/commit-msg-selftest.XXXXXX")
trap 'rm -rf "$work"' EXIT

# expect <accept|reject> <label> <message>
expect() {
  local got=accept
  printf '%s\n' "$3" > "$work/COMMIT_EDITMSG"
  bash "$hook" "$work/COMMIT_EDITMSG" > /dev/null 2>&1 || got=reject
  if [ "$got" = "$1" ]; then
    echo "  ok    $2"
  else
    echo "  FAIL  $2 (expected $1, got $got)"
    failures=$((failures + 1))
  fi
}

expect accept 'type and scope' 'fix(qa): call the reply and forward draft smokes with --comment'
expect accept 'breaking change marker' 'refactor(cli)!: remove the update command'
expect accept 'header after template comments, with a body' $'# Please enter the commit message\n\ndocs: record the shipped gates\n\nA body line.'
expect accept 'git merge header' "Merge branch 'main' into claude/admiring-brown-h2ac2z"
expect accept 'git revert header' 'Revert "feat(fs): atomic replace"'
expect accept 'git fixup header' 'fixup! fix(qa): call the smokes with --comment'
expect reject 'no type' 'update the smokes'
expect reject 'unlisted type' 'wip: half of the smokes'
expect reject 'capitalised type' 'Fix: call the smokes with --comment'
expect reject 'trailing period' 'fix(qa): call the smokes with --comment.'
expect accept 'header of exactly 100 characters' "fix(qa): $(printf 'x%.0s' $(seq 1 91))"
expect reject 'header over 100 characters' "fix(qa): $(printf 'x%.0s' $(seq 1 92))"
expect reject 'empty message' '# only a comment'

if [ "$failures" -ne 0 ]; then
  echo "commit-msg selftest: ${failures} case(s) failed" >&2
  exit 1
fi
echo "commit-msg selftest: all cases passed"
