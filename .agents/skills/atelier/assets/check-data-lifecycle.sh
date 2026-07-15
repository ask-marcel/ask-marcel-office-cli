#!/usr/bin/env bash
#
# Rule 30 tripwire: data changes are additive and reversible.
#
# Checks STAGED ADDED LINES (`--all` scans the tree for adopt-mode audits).
#
# What it catches:
#   1. A hard delete in application code:  db.delete( | deleteById( | deleteAll( | DELETE FROM
#      Exception by PATH convention, never inline: files whose path contains
#      erasure, retention, prune, or sweep (the sanctioned rule-30 exceptions:
#      privacy subject-erasure and the retention sweep).
#   2. A destructive in-place schema change in a NEW migration:
#      DROP COLUMN | RENAME COLUMN
#      Exception by NAME convention: a migration whose filename contains
#      "contract" is the deliberate contract step of expand-contract.
#
# Collection APIs (Map.delete, Set.delete, cache.delete) are not matched.
# A tripwire, not a proof (see skills/atelier/references/reliability.md).

set -euo pipefail

MODE="${1:-staged}"

HARD_DELETE='(db\.delete\(|deleteById\(|deleteAll\(|DELETE[[:space:]]+FROM)'
DESTRUCTIVE_DDL='(DROP[[:space:]]+COLUMN|RENAME[[:space:]]+COLUMN)'
EXEMPT_PATHS='erasure|retention|prune|sweep'

staged_added() { # $1 = path glob
  git diff --cached -U0 -- "$1" \
    | awk '/^\+\+\+ b\//{f=substr($0,7)} /^\+[^+]/{print f": "substr($0,2)}' || true
}

status=0

# 1. Hard deletes in application code (src/**, tests exempt, exception paths exempt).
if [ "$MODE" = "--all" ]; then
  hits=$(grep -rEn "$HARD_DELETE" --include='*.ts' --include='*.java' src/ 2>/dev/null || true)
else
  hits=$(staged_added 'src/' | grep -E "$HARD_DELETE" || true)
fi
hits=$(echo "$hits" | grep -v -E "\.test\.|test-helpers/|src/test/" | grep -v -E "$EXEMPT_PATHS" | grep -v '^$' || true)
if [ -n "$hits" ]; then
  echo "  ╳ hard delete in application code (rule 30: soft-delete by default):" >&2
  echo "$hits" | sed 's/^/    /' >&2
  status=1
fi

# 2. Destructive DDL in new migrations (contract-step files exempt by name).
if [ "$MODE" = "--all" ]; then
  ddl=$(grep -rEln "$DESTRUCTIVE_DDL" --include='*.sql' --include='*.ts' . 2>/dev/null | grep -iE 'migration' || true)
else
  ddl=$(staged_added '*migration*' | grep -E "$DESTRUCTIVE_DDL" || true)
fi
ddl=$(echo "$ddl" | grep -v -i 'contract' | grep -v '^$' || true)
if [ -n "$ddl" ]; then
  echo "  ╳ destructive schema change outside a contract-step migration (rule 30: expand-contract):" >&2
  echo "$ddl" | sed 's/^/    /' >&2
  status=1
fi

[ "$status" -eq 0 ] || echo "  fix: deletedAt stamp / expand then contract; exceptions live in erasure|retention paths or *contract* migrations (references/reliability.md)" >&2
exit "$status"
