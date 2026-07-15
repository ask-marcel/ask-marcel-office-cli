#!/usr/bin/env bash
#
# Rule 27 tripwire: personal data never in URLs, query strings, or log lines.
#
# Checks the STAGED ADDED LINES (like gitleaks protect --staged), so it blocks
# a violation entering history without flooding a brownfield tree. `--all`
# scans the whole tree instead (adopt-mode audit).
#
# What it catches (conservative, concrete patterns only):
#   1. A natural identifier as a query parameter:      ?email= / &phone= / ?ssn= / &token=
#   2. A natural identifier built into a query via URLSearchParams construction:
#      new URLSearchParams({ email })   (the incremental .set/.append form stays a review duty;
#      this is URLSearchParams-specific, so a POST body or FormData carrying email is untouched)
#   3. A natural identifier interpolated into a logger message string:
#      logger.info(`... ${user.email} ...`)   (the redactor covers meta KEYS, not message text)
#   4. Java: @QueryParam("email"|"phone"|"ssn"|"token")
#
# This is a tripwire, not a proof: it cannot see every channel (rule 27 remains
# a review duty; see skills/atelier/references/privacy.md). A hit is a hard stop;
# there is no inline suppression (rule 15). Test files are exempt.

set -euo pipefail

MODE="${1:-staged}"

QUERY_PATTERN='[?&](email|phone|ssn|token)='
URL_PARAMS_PATTERN='URLSearchParams\([^)]*(email|phone|ssn|token)'
LOG_PATTERN='logger\.(info|warn|error|debug)\([^)]*\$\{[^}]*(email|phone|ssn)'
JAVA_PATTERN='@QueryParam\("(email|phone|ssn|token)"'

added_lines() {
  if [ "$MODE" = "--all" ]; then
    grep -rEn "$1" --include='*.ts' --include='*.tsx' --include='*.java' src/ 2>/dev/null \
      | grep -v -E '\.test\.(ts|tsx)|test-helpers/|src/test/' || true
  else
    git diff --cached -U0 -- 'src/' \
      | awk '/^\+\+\+ b\//{f=substr($0,7)} /^\+[^+]/{print f": "substr($0,2)}' \
      | grep -v -E '\.test\.(ts|tsx)|test-helpers/|src/test/' \
      | grep -En "$1" || true
  fi
}

status=0
for p in "$QUERY_PATTERN" "$URL_PARAMS_PATTERN" "$LOG_PATTERN" "$JAVA_PATTERN"; do
  hits=$(added_lines "$p")
  if [ -n "$hits" ]; then
    echo "  ╳ personal data in a URL, query string, or log message (rule 27):" >&2
    echo "$hits" | sed 's/^/    /' >&2
    status=1
  fi
done

[ "$status" -eq 0 ] || echo "  fix: send personal data in a POST body; log opaque ids only (references/privacy.md)" >&2
exit "$status"
