#!/usr/bin/env bash
#
# Fast staged lint for the pre-commit hook (canon 15.1): run ESLint on the
# staged TS files only, so the hook stays O(staged files) and quick. The full,
# type-aware, zero-warning `lint:strict` runs in CI (assets/ci.yml), where its
# ~25s cost does not sit between the developer and every commit.
#
# Wire it in package.json:  "lint:staged": "bash scripts/lint-staged.sh"
#
set -euo pipefail

# NUL-separated, so a path with a space stays one argument.
files=()
while IFS= read -r -d '' f; do
  case "$f" in *.ts | *.tsx) files+=("$f") ;; esac
done < <(git diff --cached --name-only --diff-filter=ACMR -z)

if [ "${#files[@]}" -eq 0 ]; then
  echo "  no staged TS files"
  exit 0
fi

# Non-type-aware lint only (LINT_STRICT unset), so this stays fast. The
# type-aware pass is CI's job. --max-warnings=0 makes a warning block the commit
# the way it blocks CI (prettier/prettier is a warning; without the flag the hook
# passed it until 2026-09-27, rule 15's zero warnings), and --no-warn-ignored keeps
# a staged file under an ignored path from failing on ESLint's "File ignored"
# warning.
bun x eslint --max-warnings=0 --no-warn-ignored "${files[@]}"
