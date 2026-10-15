#!/usr/bin/env bash
#
# Fast staged lint for the pre-commit hook (rule 15.1): run ESLint on the
# staged TS files only, so the hook stays O(staged files) and quick. The full,
# type-aware, zero-warning `lint:strict` runs in CI (assets/ci.yml), where its
# ~25s cost does not sit between the developer and every commit.
#
# Wire it in package.json:  "lint:staged": "bash scripts/lint-staged.sh"
#
# Each folder lints its own files with its own eslint.config.js (package
# split, D7). Run from a workspace folder (`bash ../scripts/lint-staged.sh`),
# it takes the staged files under that folder; run from the root, the staged
# files outside every workspace folder.
#
set -euo pipefail

# The folder comes from where this script lives, not from git: during a hook
# in a linked worktree, git exports an absolute GIT_DIR, and git then takes
# the current directory as the top of the work tree, so `--relative` and
# `--show-prefix` see no folder.
root=$(cd "$(dirname "$0")/.." && pwd -P)
here=$(pwd -P)
prefix=${here#"$root"}
prefix=${prefix#/}

files=$(cd "$root" && git diff --cached --name-only --diff-filter=ACMR | grep -E '\.(ts|tsx)$' || true)

if [ -n "$prefix" ]; then
  files=$(printf '%s\n' "$files" | sed -n "s|^${prefix}/||p")
else
  workspaces=$(bun --print "((await Bun.file('package.json').json()).workspaces ?? []).join('|')")
  files=$(printf '%s\n' "$files" | grep -vE "^(${workspaces})/" || true)
fi

if [ -z "$files" ]; then
  echo "  no staged TS files"
  exit 0
fi

# Non-type-aware lint only (LINT_STRICT unset), so this stays fast. The
# type-aware pass is CI's job. eslint exits non-zero on any error, blocking
# the commit.
echo "$files" | xargs bun x eslint
