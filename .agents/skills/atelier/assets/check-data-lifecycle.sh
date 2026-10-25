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
#      DROP COLUMN | DROP TABLE | RENAME COLUMN | RENAME TO | TRUNCATE | ALTER COLUMN ... TYPE
#      Exception by NAME convention: a migration whose filename contains
#      "contract" is the deliberate contract step of expand-contract.
#
# Every exemption is matched against the PATH of a hit (the part before the
# first colon of a `path: content` or `path:line:content` line), never against
# the content, so a comment naming the exception cannot exempt the line it sits
# on (a `// retention` beside a hard delete, a `-- contract` beside a DROP;
# found 2026-09-02).
#
# Collection APIs (Map.delete, Set.delete, cache.delete) are not matched.
# A tripwire, not a proof (see skills/atelier/references/reliability.md).

set -euo pipefail

MODE="${1:-staged}"

# SQL is case-insensitive, so a lowercase statement is a statement (`delete from` in a
# quoted query and `drop table` in a migration passed until 2026-09-27). The DDL is
# matched case-insensitively because only migration files are read; in application
# code the lowercase form counts only as a quoted SQL string, so a comment saying
# "delete from the cache" stays prose.
HARD_DELETE='(db\.delete\(|deleteById\(|deleteAll\(|DELETE[[:space:]]+FROM|['"'"'"`][[:space:]]*[Dd][Ee][Ll][Ee][Tt][Ee][[:space:]]+[Ff][Rr][Oo][Mm][[:space:]])'
DESTRUCTIVE_DDL='(DROP[[:space:]]+(COLUMN|TABLE)|RENAME[[:space:]]+(COLUMN|TO)|TRUNCATE|ALTER[[:space:]]+COLUMN[^;]*TYPE)'
# Path-anchored: `[^:]*` stops at the first colon, so only the path is read.
EXEMPT_PATHS='^[^:]*(erasure|retention|prune|sweep)'
TEST_PATHS='^[^:]*(\.test\.|test-helpers/|src/test/)'
CONTRACT_PATHS='^[^:]*[Cc]ontract'

staged_added() { # $1 = path glob; header-aware, so an added line starting with `+` is read
  git diff --cached -U0 -- "$1" \
    | awk '/^diff --git /{h=1; next} h && /^\+\+\+ /{f=substr($0,7); next} /^@@/{h=0; next} !h && /^\+/{print f": "substr($0,2)}' || true
}

status=0

# 1. Hard deletes in application code (src/**, tests exempt, exception paths exempt).
if [ "$MODE" = "--all" ]; then
  hits=$(grep -rEn "$HARD_DELETE" --include='*.ts' --include='*.java' src/ 2>/dev/null || true)
else
  hits=$(staged_added 'src/' | grep -E "$HARD_DELETE" || true)
fi
hits=$(echo "$hits" | grep -v -E "$TEST_PATHS" | grep -v -E "$EXEMPT_PATHS" | grep -v '^$' || true)
if [ -n "$hits" ]; then
  echo "  ╳ hard delete in application code (rule 30: soft-delete by default):" >&2
  echo "$hits" | sed 's/^/    /' >&2
  status=1
fi

# 2. Destructive DDL in new migrations (contract-step files exempt by name).
if [ "$MODE" = "--all" ]; then
  ddl=$(grep -rilE "$DESTRUCTIVE_DDL" --include='*.sql' --include='*.ts' \
          --exclude-dir=node_modules --exclude-dir=.git --exclude-dir=.stryker-tmp --exclude-dir=dist --exclude-dir=build \
          . 2>/dev/null | grep -iE 'migration' || true)
else
  ddl=$(staged_added '*migration*' | grep -iE "$DESTRUCTIVE_DDL" || true)
fi
ddl=$(echo "$ddl" | grep -v -E "$CONTRACT_PATHS" | grep -v '^$' || true)
if [ -n "$ddl" ]; then
  echo "  ╳ destructive schema change outside a contract-step migration (rule 30: expand-contract):" >&2
  echo "$ddl" | sed 's/^/    /' >&2
  status=1
fi

[ "$status" -eq 0 ] || echo "  fix: deletedAt stamp / expand then contract; exceptions live in erasure|retention paths or *contract* migrations (references/reliability.md)" >&2
exit "$status"
