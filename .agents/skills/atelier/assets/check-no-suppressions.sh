#!/usr/bin/env bash
#
# Rule 15 tripwire, Java variant: no inline ignore, ever.
#
# Checks the STAGED ADDED LINES of *.java files under src/ (`--all` scans the
# tree for adopt-mode audits and for CI, where the staged index does not exist).
#
# What it catches, as text, so no tool's own suppression mechanism can hide it:
#   @SuppressWarnings(...)   javac and PMD honour it, and PMD honours it even on
#                            the rule that bans it (@SuppressWarnings("PMD")
#                            suppresses its own report), which is why the
#                            NoSuppressWarnings XPath rule in pmd-ruleset.xml is
#                            defence in depth and this grep is the authority
#   @SuppressFBWarnings      SpotBugs
#   // NOPMD                 PMD's marker; the canonical pom also sets an impossible
#                            suppressMarker so the comment is inert either way
#   // NOSONAR               SonarQube / SonarLint
#   // CHECKSTYLE:OFF        Checkstyle (and CHECKSTYLE.OFF)
#   // noinspection          IntelliJ
#
# Test code is in scope too: rule 15 has no test carve-out (a suppressed test
# is a suppressed finding). The fix is never the comment: refactor, or change
# the rule's severity at project level with a reason (references/workflow.md,
# Zero warnings; no inline ignores).
#
#   bash scripts/check-no-suppressions.sh          # staged (the fast hook)
#   bash scripts/check-no-suppressions.sh --all    # the tree (CI, adopt mode)

set -euo pipefail

MODE="${1:-staged}"

all_lines() {
  if [ "$MODE" = "--all" ]; then
    find src -type f -name '*.java' 2>/dev/null \
      | while IFS= read -r f; do awk -v p="$f" '{ printf "%s:%d: %s\n", p, NR, $0 }' "$f"; done
  else
    git diff --cached -U0 -- 'src/*.java' 'src/**/*.java' \
      | awk '/^diff --git /{h=1; next} h && /^\+\+\+ /{f=substr($0,7); next} /^@@/{h=0; next} !h && /^\+/{print f": "substr($0,2)}' || true
  fi
}

hits=$(all_lines | grep -E '@SuppressWarnings|@SuppressFBWarnings|\bNOPMD\b|\bNOSONAR\b|CHECKSTYLE[:.]OFF|noinspection' || true)

if [ -n "$hits" ]; then
  echo "  ╳ inline suppression in Java source; refactor or change the severity at project level (rule 15):" >&2
  echo "$hits" | sed 's/^/    /' >&2
  exit 1
fi
echo "  ✓ no inline suppression (rule 15)"
