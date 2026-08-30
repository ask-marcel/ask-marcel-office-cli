#!/usr/bin/env bash
#
# Fast staged lint for the pre-commit hook (rule 15.1): run ESLint on the
# staged TS files only, so the hook stays O(staged files) and quick. The full,
# type-aware, zero-warning `lint:strict` runs in CI (assets/ci.yml), where its
# ~25s cost does not sit between the developer and every commit.
#
# Wire it in package.json:  "lint:staged": "bash scripts/lint-staged.sh"
#
set -euo pipefail

files=$(git diff --cached --name-only --diff-filter=ACMR | grep -E '\.(ts|tsx)$' || true)

if [ -z "$files" ]; then
  echo "  no staged TS files"
  exit 0
fi

# Non-type-aware lint only (LINT_STRICT unset), so this stays fast. The
# type-aware pass is CI's job. eslint exits non-zero on any error, blocking
# the commit.
echo "$files" | xargs bun x eslint
