#!/usr/bin/env bash
#
# Rule 28 tripwire: every new owner-scoped endpoint ships its cross-tenant test.
#
# Heuristic: a NEW STAGED route/resource file (under the route globs below) must
# be accompanied by a test file, staged or existing, in the same directory tree
# whose content mentions 404 (the cross-tenant not_found assertion). `--all`
# audits every route file in the tree instead.
#
# CONFIG: adjust the globs to the repo's route layout.
ROUTE_GLOBS_TS='src/infra/http/'
ROUTE_GLOBS_JAVA='src/main/java/.*/api/'
#
# This is the weakest of the four guards by design: it proves a 404 test EXISTS
# near the route, not that it asserts the right thing. The real contract is the
# per-endpoint test of references/isolation.md; this wire just refuses the
# common failure of landing a route with no isolation test at all. Health and
# public routes: name them *public* or *health* to exempt them (path convention,
# no inline suppression).

set -euo pipefail

MODE="${1:-staged}"

route_files() {
  if [ "$MODE" = "--all" ]; then
    { find src/infra/http -type f -name '*.ts' 2>/dev/null;
      find src/main/java -type f -name '*.java' -path '*/api/*' 2>/dev/null; } || true
  else
    git diff --cached --name-only --diff-filter=A 2>/dev/null \
      | grep -E "^(${ROUTE_GLOBS_TS}.*\.ts|${ROUTE_GLOBS_JAVA}.*\.java)$" || true
  fi
}

has_404_test_near() { # $1 = route file
  local dir base
  dir=$(dirname "$1")
  base=$(basename "$1")
  base="${base%.*}"
  # same-dir tests, mirrored test tree (java), or staged test files
  { grep -rlE '404' "$dir" 2>/dev/null | grep -E '\.test\.(ts|tsx)$' || true;
    find "src/test" -type f -name "*${base%Resource}*" 2>/dev/null | xargs grep -lE '404' 2>/dev/null || true;
    git diff --cached --name-only 2>/dev/null | grep -E '\.test\.ts$|src/test/.*\.java$' \
      | while IFS= read -r t; do git show ":$t" 2>/dev/null | grep -qE '404' && echo "$t"; done || true; } | grep -q .
}

status=0
while IFS= read -r f; do
  [ -n "$f" ] || continue
  case "$f" in *.test.ts|*public*|*health*|*to-response*) continue ;; esac
  if ! has_404_test_near "$f"; then
    echo "  ╳ $f lands with no nearby test asserting a 404 (rule 28: cross-tenant not_found)" >&2
    status=1
  fi
done < <(route_files)

[ "$status" -eq 0 ] || echo "  fix: ship the owner-A-token-vs-owner-B-resource test with the route (references/isolation.md)" >&2
exit "$status"
