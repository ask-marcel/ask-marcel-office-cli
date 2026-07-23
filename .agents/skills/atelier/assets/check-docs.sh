#!/usr/bin/env bash
#
# docs-check (rule 12.1): run the README's documented Verify commands so a stale
# README fails CI instead of rotting until "works on my machine" is the only real
# onboarding path.
#
# It extracts the fenced ```bash blocks under the "## Verify" heading of the README
# and runs them; any non-zero command fails the check. Keep the Verify block
# self-contained and fast (a health curl, a smoke command, a path assertion), not
# the full install, so it is cheap enough to run on every pull request.
#
# Wire it as a docs-check CI job (see references/governance.md); prove it can fail
# by pointing it at a README whose documented command no longer works.
#
#   bash scripts/check-docs.sh [README.md]
#
set -euo pipefail

README="${1:-README.md}"
if [ ! -f "$README" ]; then
  echo "docs-check: $README not found" >&2
  exit 1
fi

# Capture the lines inside ```bash fences that sit under the "## Verify" heading.
block=$(awk '
  /^## / { in_verify = ($0 ~ /^##[[:space:]]+Verify([[:space:]]|$)/) }
  in_verify && /^```bash[[:space:]]*$/ { in_fence = 1; next }
  in_verify && /^```/ { in_fence = 0; next }
  in_verify && in_fence { print }
' "$README")

if [ -z "$block" ]; then
  echo "docs-check: no '## Verify' bash block in $README, nothing to run" >&2
  exit 0
fi

echo "docs-check: running the README Verify commands..."
# -e so the first failing command fails the check; the README is the source of truth.
bash -eu -c "$block"
echo "docs-check: README Verify commands passed"
