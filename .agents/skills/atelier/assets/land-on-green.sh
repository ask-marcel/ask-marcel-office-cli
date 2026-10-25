#!/usr/bin/env bash
#
# Land a pull request once CI is green: rebase merge, the branch deleted, pinned to the head CI
# tested (rule 38). The `land` job of the shipped CI workflow runs it after every other job of that
# workflow passed, so nobody presses merge and nothing lands red (references/workflow.md, Landing on
# green).
#
# It refuses, red, when:
#   - another check on the head is still running after LAND_WAIT_MIN minutes (default 20), or
#     finished other than success, skipped or neutral (a dependency audit in its own workflow, a
#     review bot): `needs:` sees only this workflow's jobs;
#   - the base branch moved since CI ran. CI tested the merge of the head onto the base as it was
#     then (the event's merge commit, GITHUB_SHA, whose first parent is that base), and a merge
#     made with GITHUB_TOKEN starts no push run on main, so landing onto a newer base would put an
#     untested tree on main. Rebase onto the new main and push: CI runs again and lands it.
#
# A draft pull request, one from a fork and a repository variable LAND_ON_GREEN=off never reach
# this script (the job's `if:`). Inputs, all set by the workflow: GH_TOKEN, GITHUB_REPOSITORY,
# GITHUB_SHA, PR_NUMBER, PR_HEAD_SHA, PR_BASE_REF, and LAND_JOB (this job's check name).
#
# Usage:
#   bash scripts/land-on-green.sh              # in the land job
#   bash scripts/land-on-green.sh --selftest   # prove it refuses, with a stub gh
set -euo pipefail

land() {
  local repo="$GITHUB_REPOSITORY" pr="$PR_NUMBER" head="$PR_HEAD_SHA" base_ref="$PR_BASE_REF"
  local self="${LAND_JOB:-land}" wait_min="${LAND_WAIT_MIN:-20}" poll="${LAND_POLL_SEC:-20}"

  # The merge commit CI checked out: parent 1 is the base it tested, parent 2 the head.
  local parents tested_base tested_head
  parents=$(gh api "repos/$repo/commits/$GITHUB_SHA" --jq '[.parents[].sha] | join(" ")')
  tested_base=${parents%% *}
  tested_head=${parents##* }
  if [ "$tested_head" != "$head" ]; then
    echo "land: CI tested head ${tested_head:0:7}, but the pull request's head is ${head:0:7}; the newer run lands it" >&2
    return 1
  fi

  # Every other check on the head: wait while one runs, then all must have passed.
  local deadline=$(( $(date +%s) + wait_min * 60 )) runs pending failed
  while :; do
    runs=$(gh api "repos/$repo/commits/$head/check-runs?per_page=100" --paginate \
      --jq '.check_runs[] | [.name, .status, (.conclusion // "")] | @tsv')
    runs=$(printf '%s\n' "$runs" | awk -F'\t' -v self="$self" 'NF && $1 != self')
    pending=$(printf '%s\n' "$runs" | awk -F'\t' 'NF && $2 != "completed" {print $1}')
    [ -z "$pending" ] && break
    if [ "$(date +%s)" -ge "$deadline" ]; then
      echo "land: still running after ${wait_min} min, not landing:" >&2
      printf '%s\n' "$pending" | sed 's/^/    /' >&2
      return 1
    fi
    sleep "$poll"
  done
  failed=$(printf '%s\n' "$runs" | awk -F'\t' 'NF && $3 != "success" && $3 != "skipped" && $3 != "neutral" {print $1 " (" $3 ")"}')
  if [ -n "$failed" ]; then
    echo "land: a check on the head did not pass, not landing:" >&2
    printf '%s\n' "$failed" | sed 's/^/    /' >&2
    return 1
  fi

  # The base must still be the one CI tested.
  local current
  current=$(gh api "repos/$repo/git/ref/heads/$base_ref" --jq '.object.sha')
  if [ "$current" != "$tested_base" ]; then
    echo "land: $base_ref moved since CI ran (tested ${tested_base:0:7}, now ${current:0:7}); rebase onto $base_ref and push, and CI lands it" >&2
    return 1
  fi

  echo "land: CI green on ${head:0:7} over $base_ref ${current:0:7}; rebase merge, branch deleted"
  gh pr merge "$pr" --repo "$repo" --rebase --delete-branch --match-head-commit "$head"
}

selftest() {
  local tmp fails=0
  tmp=$(mktemp -d)
  trap 'rm -rf "$tmp"' RETURN
  mkdir -p "$tmp/bin"
  # The stub answers from STUB_* variables and records a merge in $tmp/merged.
  cat > "$tmp/bin/gh" <<'STUB'
#!/usr/bin/env bash
case "$*" in
  "api repos/o/r/commits/m1 "*) echo "$STUB_PARENTS" ;;
  "api repos/o/r/commits/h1/check-runs"*) printf '%b' "$STUB_RUNS" ;;
  "api repos/o/r/git/ref/heads/main "*) echo "$STUB_MAIN" ;;
  "pr merge 7 --repo o/r --rebase --delete-branch --match-head-commit h1") echo merged > "$STUB_DIR/merged" ;;
  *) echo "stub gh: unexpected call: $*" >&2; exit 9 ;;
esac
STUB
  chmod +x "$tmp/bin/gh"

  # name | parents | check runs | main now | expect: merged or the refusal it must print
  local cases=(
    "green|b1 h1|land\tin_progress\t\ngates\tcompleted\tsuccess\naudit\tcompleted\tskipped\n|b1|merged"
    "a check failed|b1 h1|gates\tcompleted\tsuccess\naudit\tcompleted\tfailure\n|b1|did not pass"
    "a check cancelled|b1 h1|gates\tcompleted\tcancelled\n|b1|did not pass"
    "main moved|b1 h1|gates\tcompleted\tsuccess\n|b2|moved since CI ran"
    "a newer head|b1 h0|gates\tcompleted\tsuccess\n|b1|the newer run lands it"
    "a check still running|b1 h1|audit\tqueued\t\n|b1|still running"
  )
  local c name parents runs main want out
  for c in "${cases[@]}"; do
    IFS='|' read -r name parents runs main want <<<"$c"
    rm -f "$tmp/merged"
    out=$(PATH="$tmp/bin:$PATH" STUB_DIR="$tmp" STUB_PARENTS="$parents" STUB_RUNS="$runs" STUB_MAIN="$main" \
      GITHUB_REPOSITORY=o/r GITHUB_SHA=m1 PR_NUMBER=7 PR_HEAD_SHA=h1 PR_BASE_REF=main LAND_JOB=land \
      LAND_WAIT_MIN=0 LAND_POLL_SEC=0 land 2>&1) && st=0 || st=$?
    if [ "$want" = merged ]; then
      if [ "$st" -eq 0 ] && [ -f "$tmp/merged" ]; then echo "  ok:   $name lands"
      else echo "  FAIL: $name did not land: $out"; fails=$((fails + 1)); fi
    elif [ "$st" -ne 0 ] && [ ! -f "$tmp/merged" ] && grep -q "$want" <<<"$out"; then
      echo "  ok:   $name refuses ($want)"
    else
      echo "  FAIL: $name: exit $st, merged=$([ -f "$tmp/merged" ] && echo yes || echo no): $out"; fails=$((fails + 1))
    fi
  done
  [ "$fails" -eq 0 ] && echo "land-on-green: selftest passed" && return 0
  echo "land-on-green: selftest FAILED ($fails)"; return 1
}

if [ "${1:-}" = "--selftest" ]; then selftest; else land; fi
