#!/usr/bin/env bash
#
# The discipline tripwires that are core gates: rules 27 (personal data never
# in a query string or a log line), 29 (a deadline on every outbound call) and
# 30 (no hard delete, no destructive DDL outside a contract migration), run in
# order so one hook step and one CI step carry all three. Each is inert in a
# repo without the concern and wrong in any repo with it, which is why they are
# default gates. The isolation guard (rule 28, check-isolation-tests.sh) is
# NOT here: it demands a cross-tenant 404 test of every new route, which is
# wrong for a single-user app, so it is wired beside this script only where
# tenants or owners exist (references/workflow.md, Discipline tripwires).
#
# Every guard runs even after one fails, so a commit shows all its findings at
# once; the exit is non-zero if any failed. `--all` is passed through.
#
#   bash scripts/check-disciplines.sh          # staged added lines (the hook)
#   bash scripts/check-disciplines.sh --all    # the whole tree (CI, adopt audit)

set -uo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
status=0
for guard in check-pii-channels check-io-deadlines check-data-lifecycle; do
  if [ ! -f "$here/$guard.sh" ]; then
    echo "check-disciplines: $here/$guard.sh is missing; copy it from the skill's assets/" >&2
    status=1
    continue
  fi
  bash "$here/$guard.sh" "$@" || status=1
done
exit "$status"
