#!/usr/bin/env bash
#
# atelier commit-message CI gate - re-check Conventional Commits server-side
#
# The `commit-msg` hook (assets/commit-msg) is the local first line, and
# `git commit --no-verify` walks straight past it. This re-runs the identical
# grammar over every commit in the pushed range, on the one line that cannot be
# skipped, exactly as the CI gate set backstops the fast pre-commit hook
# (rule 23; canon 1.3 "One grammar for the history", which asks for the linter
# in the fast hook AND a re-check in CI).
#
# One grammar, one validator: this delegates to the hook script instead of
# restating the pattern, so the local gate and the CI gate cannot drift apart.
#
# Usage:
#   bash scripts/check-commit-messages.sh [range]
#
# The range defaults to the GitHub Actions context (the PR's base branch), then
# to origin/main..HEAD locally. Pass one explicitly to check any other span.
# Merge commits are excluded; the hook's own exemptions (Revert, fixup!,
# squash!, amend!) still apply.
#
# Installed alongside the other gate scripts:
#
#   cp <skill>/assets/check-commit-messages.sh scripts/check-commit-messages.sh
#
# See skills/atelier/references/workflow.md (Commit message format).

set -euo pipefail

hook="${COMMIT_MSG_HOOK:-.githooks/commit-msg}"

if [ ! -f "$hook" ]; then
  cat <<EOF >&2
  ╳ COMMIT-MESSAGE HOOK NOT FOUND

  Expected the validator at: ${hook}

  This gate re-runs the hook's grammar in CI. Install the hook first:
      cp <skill>/assets/commit-msg .githooks/commit-msg
      chmod +x .githooks/commit-msg
      git config core.hooksPath .githooks
EOF
  exit 1
fi

range="${1:-}"
if [ -z "$range" ]; then
  if [ -n "${GITHUB_BASE_REF:-}" ]; then
    git fetch --quiet origin "$GITHUB_BASE_REF" 2>/dev/null || true
    range="origin/${GITHUB_BASE_REF}..HEAD"
  else
    range="origin/main..HEAD"
  fi
fi

# A left side that does not resolve (a fresh branch, a shallow clone, the zero
# SHA GitHub sends for a new branch) narrows to the tip commit rather than
# failing the build on a git plumbing error. Never widen to the whole history:
# a legacy repo adopting the standard would fail on commits nobody can rewrite.
rev_args=(--no-merges "$range")
if ! git rev-parse --quiet --verify "${range%%..*}^{commit}" >/dev/null 2>&1; then
  if git rev-parse --quiet --verify 'HEAD~1^{commit}' >/dev/null 2>&1; then
    rev_args=(--no-merges HEAD~1..HEAD)
  else
    rev_args=(--no-merges --max-count=1 HEAD)
  fi
  echo "check-commit-messages: '${range}' does not resolve, checking ${rev_args[*]} instead" >&2
fi

commits=$(git rev-list "${rev_args[@]}")

if [ -z "$commits" ]; then
  echo "  ✓ commit messages: no commits in range (${rev_args[*]})"
  exit 0
fi

msg_file=$(mktemp)
trap 'rm -f "$msg_file"' EXIT

failed=0
count=0
for sha in $commits; do
  count=$((count + 1))
  git log -1 --format=%B "$sha" >"$msg_file"
  if ! output=$(bash "$hook" "$msg_file" 2>&1); then
    failed=1
    printf '\n  ╳ %s  %s\n' "$(git log -1 --format=%h "$sha")" "$(git log -1 --format=%s "$sha")" >&2
    printf '%s\n' "$output" >&2
  fi
done

if [ "$failed" -ne 0 ]; then
  cat <<EOF >&2

  ${count} commit message(s) checked; at least one is not Conventional.

  A hook bypassed with --no-verify still has to pass here. Rewrite the offending
  message (git commit --amend, or an interactive rebase for an older commit) and
  push again.
EOF
  exit 1
fi

echo "  ✓ commit messages: ${count} checked, all Conventional"
