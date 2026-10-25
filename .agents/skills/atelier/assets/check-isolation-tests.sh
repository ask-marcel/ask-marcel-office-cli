#!/usr/bin/env bash
#
# Rule 28 tripwire: every new owner-scoped endpoint ships its cross-tenant test.
#
# Heuristic: a NEW STAGED route/resource file (under the route globs below) must
# be accompanied by a test file, staged or existing, named for the route (the
# route's basename appears in the test's filename: invoices.ts and
# invoices.test.ts, InvoiceResource.java and InvoiceResourceTest.java) whose
# body asserts a 404 inside a test block (a non-comment line carrying 404
# after a `test(` / `it(` / `@Test` marker). `--all` audits every route file in
# the tree instead.
#
# Before 2026-09-02 any test file in the directory containing the characters
# 404, in a comment or a TODO, satisfied it; the stricter shape still proves
# only that a named test asserts 404 somewhere, not that it asserts the right
# thing.
#
# CONFIG: adjust the globs to the repo's route layout.
ROUTE_GLOBS_TS='src/infra/http/'
ROUTE_GLOBS_JAVA='src/main/java/.*/api/'
#
# This is the weakest of the five guards by design. The real contract is the
# per-endpoint test of references/isolation.md; this wire just refuses the
# common failure of landing a route with no isolation test at all. Health and
# public routes: name them *public* or *health* to exempt them (path convention,
# no inline suppression).

set -euo pipefail

MODE="${1:-staged}"

# Both modes read the same globs: --all (the CI mode) hardcoded src/infra/http and
# */api/* until 2026-09-27, so a repo that adjusted the globs was checked on commit
# and never in CI.
ROUTE_PATTERN="^(${ROUTE_GLOBS_TS}.*\.ts|${ROUTE_GLOBS_JAVA}.*\.java)$"
route_files() {
  if [ "$MODE" = "--all" ]; then
    find src -type f \( -name '*.ts' -o -name '*.java' \) 2>/dev/null | grep -E "$ROUTE_PATTERN" || true
  else
    git diff --cached --name-only --diff-filter=A 2>/dev/null | grep -E "$ROUTE_PATTERN" || true
  fi
}

asserts_404() { # stdin = a test file's content; exit 0 when a test block carries a 404
  awk '
    /^[[:space:]]*(\/\/|\*|\/\*)/ { next }
    { sub(/[[:space:]]*\/\/.*$/, "") }
    /(^|[^A-Za-z_])(test|it)\(|@Test/ { inblock = 1 }
    inblock && /404/ { found = 1 }
    END { exit found ? 0 : 1 }'
}

has_404_test_near() { # $1 = route file
  local dir base stem t
  dir=$(dirname "$1")
  base=$(basename "$1")
  base="${base%.*}"
  stem="${base%Resource}"
  # same-dir tests and the mirrored Java test tree, named for the route
  { find "$dir" -type f \( -name "*${base}*.test.ts" -o -name "*${base}*.test.tsx" \) 2>/dev/null;
    find src/test -type f -name "*${stem}*.java" 2>/dev/null; } | sort -u \
    | while IFS= read -r t; do [ -f "$t" ] && asserts_404 < "$t" && echo "$t"; done | grep -q . && return 0
  # staged test files named for the route (not yet on disk under a commit)
  git diff --cached --name-only 2>/dev/null | grep -E "(${base}[^/]*\.test\.tsx?|${stem}[^/]*\.java)$" \
    | while IFS= read -r t; do git show ":$t" 2>/dev/null | asserts_404 && echo "$t"; done | grep -q .
}

status=0
while IFS= read -r f; do
  [ -n "$f" ] || continue
  case "$f" in *.test.ts|*public*|*health*|*to-response*) continue ;; esac
  if ! has_404_test_near "$f"; then
    echo "  ╳ $f lands with no test named for it that asserts a 404 (rule 28: cross-tenant not_found)" >&2
    status=1
  fi
done < <(route_files)

[ "$status" -eq 0 ] || echo "  fix: ship the owner-A-token-vs-owner-B-resource test with the route (references/isolation.md)" >&2
exit "$status"
