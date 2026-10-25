#!/usr/bin/env bash
#
# Rule 27 tripwire: personal data never in URLs, query strings, or log lines.
#
# Checks the STAGED ADDED LINES (like the staged secret scan), so it blocks
# a violation entering history without flooding a brownfield tree. `--all`
# scans the whole tree instead (adopt-mode audit).
#
# What it catches (conservative, concrete patterns only), with NAMES standing
# for email | phone | ssn | token | password | secret | iban | card | dob |
# birthdate | address | firstName | lastName, matched case-insensitively and
# with any prefix (userEmail, user_email):
#   1. A natural identifier as a query parameter:      ?email= / &userPhone= / ?ssn= / &token=
#   2. A natural identifier built into a query via URLSearchParams construction:
#      new URLSearchParams({ email })   (the incremental .set/.append form stays a review duty;
#      this is URLSearchParams-specific, so a POST body or FormData carrying email is untouched)
#   3. A natural identifier interpolated into a logger message string:
#      logger.info(`... ${user.email} ...`)   (the redactor covers meta KEYS, not message text).
#      A call split over lines is read too: the call line is joined with up to
#      three following lines before matching (a one-line-only check passed
#      every multi-line logger call; found 2026-09-02).
#   4. Java: @QueryParam("email" | "customerPhone" | ...)
#
# This is a tripwire, not a proof: it cannot see every channel (rule 27 remains
# a review duty; see skills/atelier/references/privacy.md). A hit is a hard stop;
# there is no inline suppression (rule 15). Test files are exempt.

set -euo pipefail

MODE="${1:-staged}"

# One stream of `path: content` lines: the staged added lines, or every line of
# every source file under --all. The patterns then run in a single awk pass.
all_lines() {
  if [ "$MODE" = "--all" ]; then
    find src -type f \( -name '*.ts' -o -name '*.tsx' -o -name '*.java' \) 2>/dev/null \
      | grep -v -E '\.test\.(ts|tsx)|test-helpers/|src/test/' \
      | while IFS= read -r f; do awk -v p="$f" '{ print p ": " $0 }' "$f"; done
  else
    git diff --cached -U0 -- 'src/' \
      | awk '/^diff --git /{h=1; next} h && /^\+\+\+ /{f=substr($0,7); next} /^@@/{h=0; next} !h && /^\+/{print f": "substr($0,2)}' \
      | grep -v -E '^[^:]*(\.test\.(ts|tsx)|test-helpers/|src/test/)' || true
  fi
}

# The patterns live inside the awk program as literals: `awk -v` rewrites
# backslash escapes (BSD awk turns \( into a bare paren), which silently
# breaks a regex passed that way. Lines are lowercased before matching.
hits=$(all_lines | awk '
  { line[++n] = $0; file[n] = $0; sub(/: .*$/, "", file[n]) }
  END {
    names = "(email|phone|ssn|token|password|secret|iban|card|dob|birthdate|address|firstname|lastname)"
    query = "[?&][a-z_]*" names "="
    params = "urlsearchparams\\([^)]*" names
    logmsg = "logger\\.(info|warn|error|debug)\\([^)]*\\$\\{[^}]*" names
    javaq = "@queryparam\\(\"[a-z_]*" names "\""
    for (i = 1; i <= n; i++) {
      low = tolower(line[i])
      if (low ~ query || low ~ params || low ~ javaq) { print line[i]; continue }
      if (low ~ /logger\.(info|warn|error|debug)\(/) {
        joined = low
        for (k = i + 1; k <= i + 3 && k <= n && file[k] == file[i]; k++) joined = joined " " tolower(line[k])
        if (joined ~ logmsg) print line[i]
      }
    }
  }')

if [ -n "$hits" ]; then
  echo "  ╳ personal data in a URL, query string, or log message (rule 27):" >&2
  echo "$hits" | sed 's/^/    /' >&2
  echo "  fix: send personal data in a POST body; log opaque ids only (references/privacy.md)" >&2
  exit 1
fi
exit 0
