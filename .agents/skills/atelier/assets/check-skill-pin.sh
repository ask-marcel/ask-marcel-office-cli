#!/usr/bin/env bash
#
# Gate: the vendored copy of the standard must not be silently stale.
#
# A repo that vendors or pins this skill has taken a dependency on doctrine,
# and it goes stale exactly like a library does: quietly, while every other
# gate stays green. The 2026-08-30 field test found a consumer running a
# 49-day-old copy, so its hook still ran the retired eight-gate design and its
# rule 26 read the superseded wording. Nothing was broken; the repo was
# faithfully following a standard that had moved.
#
# Like the CVE scan, this belongs on a schedule and not on a commit: upstream
# doctrine changes independently of your diff, so blocking an unrelated commit
# because a skill landed overnight fails in the wrong place (references/
# workflow.md, Dependency scanning).
#
# Usage:
#   bash scripts/check-skill-pin.sh [path-to-vendored-SKILL.md]
#     default path: .claude/skills/atelier/SKILL.md
#   SKILL_PIN_UPSTREAM   raw URL or local path of the upstream SKILL.md
#   SKILL_PIN_SELFTEST   internal, used by --selftest
#
# Degrades gracefully: no vendored copy, or no network, exits 0 with a note.
# Being unable to check is not the same as being current, so the note says so.
set -euo pipefail

UPSTREAM="${SKILL_PIN_UPSTREAM:-https://raw.githubusercontent.com/vdelacou/atelier/main/skills/atelier/SKILL.md}"

hash_of() { shasum -a 256 "$1" 2>/dev/null | cut -d' ' -f1; }

# Compare two directory trees file by file. A vendored standard is a tree: the
# references carry the detail SKILL.md points at, and the assets ARE the gates,
# so a check that reads only SKILL.md can report "current" over a stale
# reference. That is the gate-that-cannot-fail shape (canon 15.10) in the gate
# meant to catch staleness, so it is checked here.
compare_trees() {
  local vend="$1" up="$2" stale=0 rel
  while IFS= read -r rel; do
    if [ ! -f "$vend/$rel" ]; then
      echo "      missing locally: $rel" >&2
      stale=$((stale + 1))
    elif [ "$(hash_of "$vend/$rel")" != "$(hash_of "$up/$rel")" ]; then
      echo "      behind upstream: $rel" >&2
      stale=$((stale + 1))
    fi
  done < <(cd "$up" && find . -type f ! -path '*/.git/*' | sed 's|^\./||' | sort)
  [ "$stale" -eq 0 ] && return 0
  echo "      $stale file(s) behind" >&2
  return 1
}

check_pin() {
  local vendored="$1" upstream="$2" tmp
  # Tree mode: both sides are directories.
  if [ -d "$vendored" ] && [ -d "$upstream" ]; then
    if compare_trees "$vendored" "$upstream" 2>/tmp/.pin-stale; then
      echo "check-skill-pin: the vendored standard matches upstream (whole tree)"
      rm -f /tmp/.pin-stale
      return 0
    fi
    {
      echo "  ╳ the vendored standard is behind upstream."
      echo ""
      cat /tmp/.pin-stale
      echo ""
      echo "  A pinned standard is a dependency (references/governance.md). Re-sync"
      echo "  deliberately: read what changed, then bring the doctrine AND the assets"
      echo "  that enforce it across in one commit. Gates and prose move together."
    } >&2
    rm -f /tmp/.pin-stale
    return 1
  fi
  if [ -d "$vendored" ] && [ ! -d "$upstream" ]; then
    # A directory vendored against a single upstream file can only check that
    # file, so say what is NOT covered rather than implying the tree is current.
    echo "check-skill-pin: comparing SKILL.md only; references and assets are not covered" >&2
    vendored="$vendored/SKILL.md"
  fi
  if [ ! -f "$vendored" ]; then
    echo "check-skill-pin: no vendored copy at $vendored, nothing to compare" >&2
    return 0
  fi
  tmp=$(mktemp)
  # shellcheck disable=SC2064
  trap "rm -f '$tmp'" RETURN
  if [ -f "$upstream" ]; then
    cp "$upstream" "$tmp"
  elif ! curl -sSfL "$upstream" -o "$tmp" 2>/dev/null; then
    echo "check-skill-pin: could not reach $upstream; the pin is UNVERIFIED, not current" >&2
    return 0
  fi
  if [ "$(hash_of "$vendored")" = "$(hash_of "$tmp")" ]; then
    echo "check-skill-pin: the vendored standard matches upstream"
    return 0
  fi
  {
    echo "  ╳ the vendored standard is behind upstream."
    echo ""
    echo "      vendored: $vendored"
    echo "      upstream: $upstream"
    echo "      differing lines: $(diff "$vendored" "$tmp" | grep -c '^[<>]' || true)"
    echo ""
    echo "  A pinned standard is a dependency (references/governance.md). Re-sync"
    echo "  deliberately: read what changed, then bring the doctrine AND the assets"
    echo "  that enforce it across in one commit. Gates and prose move together."
  } >&2
  return 1
}

selftest() {
  local tmp; tmp=$(mktemp -d)
  trap "rm -rf '$tmp'" EXIT
  printf 'doctrine v1\n' > "$tmp/upstream.md"
  printf 'doctrine v1\n' > "$tmp/current.md"
  printf 'doctrine v0 (stale)\n' > "$tmp/stale.md"
  if ! check_pin "$tmp/current.md" "$tmp/upstream.md" >/dev/null; then
    echo "selftest FAIL: an up-to-date copy was rejected" >&2; exit 1
  fi
  if check_pin "$tmp/stale.md" "$tmp/upstream.md" 2>/dev/null; then
    echo "selftest FAIL: a stale copy was accepted" >&2; exit 1
  fi
  if ! check_pin "$tmp/absent.md" "$tmp/upstream.md" >/dev/null 2>&1; then
    echo "selftest FAIL: a repo with no vendored copy should pass" >&2; exit 1
  fi
  if ! check_pin "$tmp/current.md" "http://127.0.0.1:9/unreachable" >/dev/null 2>&1; then
    echo "selftest FAIL: an unreachable upstream should degrade, not block" >&2; exit 1
  fi

  # A vendored standard is a TREE. SKILL.md matching upstream says nothing about
  # the references and assets beside it, and a gate that reports "current" while
  # a reference is stale is worse than no gate (found in a real consumer repo,
  # 2026-08-30: SKILL.md current, references/java-quarkus.md behind).
  mkdir -p "$tmp/up/references" "$tmp/vend/references"
  printf 'doctrine v1\n' > "$tmp/up/SKILL.md"
  printf 'detail v2\n'   > "$tmp/up/references/x.md"
  printf 'doctrine v1\n' > "$tmp/vend/SKILL.md"
  printf 'detail v1\n'   > "$tmp/vend/references/x.md"
  if check_pin "$tmp/vend" "$tmp/up" >/dev/null 2>&1; then
    echo "selftest FAIL: a stale reference passed while SKILL.md matched" >&2; exit 1
  fi
  printf 'detail v2\n' > "$tmp/vend/references/x.md"
  if ! check_pin "$tmp/vend" "$tmp/up" >/dev/null; then
    echo "selftest FAIL: a fully current tree was rejected" >&2; exit 1
  fi
  echo "selftest OK: gate rejects a stale vendored copy, accepts a current one, degrades when it cannot check"
}

if [ "${1:-}" = "--selftest" ]; then
  selftest
  exit 0
fi

check_pin "${1:-.claude/skills/atelier/SKILL.md}" "$UPSTREAM"
