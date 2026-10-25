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
#   bash scripts/check-skill-pin.sh [path-to-vendored-skill]
#     default: .claude/skills/atelier, the whole tree (SKILL.md, references,
#     assets), compared file by file. A vendored standard is a tree: the
#     references carry the detail SKILL.md points at and the assets ARE the
#     gates, so a check that reads only SKILL.md can report "current" over a
#     stale reference (the gate-that-cannot-fail shape, canon 15.10; seen on
#     a real consumer 2026-08-30).
#   SKILL_PIN_UPSTREAM   where the current standard lives, one of:
#     - a local directory: the skill dir of a checkout
#     - a git repository URL (https://..., git@..., file://...), cloned shallowly;
#       SKILL_PIN_SUBDIR (default skills/atelier) is the skill dir inside it
#     - a single file, local path or raw URL ending in .md: SKILL.md only, and
#       the report says the references and assets are then NOT covered
#     The shipped audit workflow sets it to the public repository; point it at
#     a fork if you keep one. Unset, the check cannot run and says so.
#   SKILL_PIN_SUBDIR     the skill dir inside a cloned repository (default skills/atelier)
#
# Degrades gracefully: no vendored copy, no upstream configured, or no network,
# exits 0 with a note. Being unable to check is not the same as being current,
# so the note says so.
set -euo pipefail

UPSTREAM="${SKILL_PIN_UPSTREAM:-}"
SUBDIR="${SKILL_PIN_SUBDIR:-skills/atelier}"
ZERO_SHA_NOTE="the pin is UNVERIFIED, not current"

WORK=$(mktemp -d)
trap 'rm -rf "$WORK"' EXIT

# Byte comparison, no hasher: `shasum` missing from PATH made both sides an empty
# string, and two empty strings are equal, so a stale tree read as current until
# 2026-09-27. cmp exits 0 same, 1 different, 2 on an error, and only 0 is a match.
same_file() { cmp -s "$1" "$2"; }

# Compare two directory trees file by file; the list of stale files goes to
# stderr, the count is the return value's message.
compare_trees() {
  local vend="$1" up="$2" stale=0 seen=0 rel
  while IFS= read -r rel; do
    seen=$((seen + 1))
    if [ ! -f "$vend/$rel" ]; then
      echo "      missing locally: $rel" >&2
      stale=$((stale + 1))
    elif ! same_file "$vend/$rel" "$up/$rel"; then
      echo "      behind upstream: $rel" >&2
      stale=$((stale + 1))
    fi
  done < <(cd "$up" && find . -type f ! -path '*/.git/*' | sed 's|^\./||' | sort)
  # An upstream listing that yields nothing (find missing, an unreadable clone) is a
  # failure to check, never a match: zero files compared is zero files current.
  if [ "$seen" -eq 0 ]; then
    echo "      the upstream tree listed no files ($up); nothing was compared" >&2
    return 1
  fi
  [ "$stale" -eq 0 ] && return 0
  echo "      $stale file(s) behind" >&2
  return 1
}

resync_advice() {
  echo "  A pinned standard is a dependency (references/governance.md). Re-sync" >&2
  echo "  deliberately: read what changed, then bring the doctrine AND the assets" >&2
  echo "  that enforce it across in one commit. Gates and prose move together." >&2
}

# Turn SKILL_PIN_UPSTREAM into a local path to compare against (stdout).
# Returns 1 when it cannot be reached or does not have the expected shape.
resolve_upstream() {
  local src="$1"
  if [ -d "$src" ] || [ -f "$src" ]; then
    echo "$src"
  elif [[ "$src" =~ ^https?:// ]] && [[ "$src" =~ \.md$ ]]; then
    curl -sSfL "$src" -o "$WORK/upstream.md" 2>/dev/null || return 1
    echo "$WORK/upstream.md"
  elif [[ "$src" =~ ^(https?://|git@|ssh://|file://) ]]; then
    # Conversion off: compare upstream's committed bytes, not the machine's line endings.
    # Git for Windows sets core.autocrlf=true system-wide, and a CRLF checkout made every
    # file of a current LF copy read as behind (2026-09-28).
    git -c core.autocrlf=false clone --quiet --depth 1 "$src" "$WORK/clone" 2>/dev/null || return 1
    [ -d "$WORK/clone/$SUBDIR" ] || { echo "check-skill-pin: $src has no $SUBDIR directory" >&2; return 1; }
    echo "$WORK/clone/$SUBDIR"
  else
    return 1
  fi
}

check_pin() {
  local vendored="$1" upstream="$2"
  # Tree mode: both sides are directories.
  if [ -d "$vendored" ] && [ -d "$upstream" ]; then
    if compare_trees "$vendored" "$upstream" 2>"$WORK/stale"; then
      echo "check-skill-pin: the vendored standard matches upstream (whole tree)"
      return 0
    fi
    { echo "  ╳ the vendored standard is behind upstream."; echo ""; cat "$WORK/stale"; echo ""; } >&2
    resync_advice
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
  if [ -d "$upstream" ]; then
    upstream="$upstream/SKILL.md"
  fi
  if same_file "$vendored" "$upstream"; then
    echo "check-skill-pin: the vendored standard matches upstream"
    return 0
  fi
  {
    echo "  ╳ the vendored standard is behind upstream."
    echo ""
    echo "      vendored: $vendored"
    echo "      upstream: $upstream"
    echo "      differing lines: $(diff "$vendored" "$upstream" | grep -c '^[<>]' || true)"
    echo ""
  } >&2
  resync_advice
  return 1
}

run() {
  local vendored="$1" up
  if [ ! -e "$vendored" ]; then
    echo "check-skill-pin: no vendored copy at $vendored, nothing to compare" >&2
    return 0
  fi
  if [ -z "$UPSTREAM" ]; then
    echo "check-skill-pin: SKILL_PIN_UPSTREAM is not set; $ZERO_SHA_NOTE" >&2
    return 0
  fi
  if ! up=$(resolve_upstream "$UPSTREAM"); then
    echo "check-skill-pin: could not reach $UPSTREAM; $ZERO_SHA_NOTE" >&2
    return 0
  fi
  check_pin "$vendored" "$up"
}

selftest() {
  local t="$WORK/selftest"
  mkdir -p "$t"
  printf 'doctrine v1\n' > "$t/upstream.md"
  printf 'doctrine v1\n' > "$t/current.md"
  printf 'doctrine v0 (stale)\n' > "$t/stale.md"
  if ! check_pin "$t/current.md" "$t/upstream.md" >/dev/null; then
    echo "selftest FAIL: an up-to-date copy was rejected" >&2; exit 1
  fi
  if check_pin "$t/stale.md" "$t/upstream.md" 2>/dev/null; then
    echo "selftest FAIL: a stale copy was accepted" >&2; exit 1
  fi
  if ! UPSTREAM="$t/upstream.md" run "$t/absent.md" >/dev/null 2>&1; then
    echo "selftest FAIL: a repo with no vendored copy should pass" >&2; exit 1
  fi
  if ! UPSTREAM="http://127.0.0.1:9/unreachable.md" run "$t/current.md" >/dev/null 2>&1; then
    echo "selftest FAIL: an unreachable upstream should degrade, not block" >&2; exit 1
  fi
  if ! UPSTREAM="" run "$t/current.md" >/dev/null 2>&1; then
    echo "selftest FAIL: an unset upstream should degrade, not block" >&2; exit 1
  fi

  # Tree mode: SKILL.md matching upstream says nothing about the references and
  # assets beside it (a real consumer, 2026-08-30: SKILL.md current,
  # references/java-quarkus.md behind).
  mkdir -p "$t/up/references" "$t/vend/references"
  printf 'doctrine v1\n' > "$t/up/SKILL.md"
  printf 'detail v2\n'   > "$t/up/references/x.md"
  printf 'doctrine v1\n' > "$t/vend/SKILL.md"
  printf 'detail v1\n'   > "$t/vend/references/x.md"
  if check_pin "$t/vend" "$t/up" >/dev/null 2>&1; then
    echo "selftest FAIL: a stale reference passed while SKILL.md matched" >&2; exit 1
  fi
  printf 'detail v2\n' > "$t/vend/references/x.md"
  if ! check_pin "$t/vend" "$t/up" >/dev/null; then
    echo "selftest FAIL: a fully current tree was rejected" >&2; exit 1
  fi
  # Zero files compared is not a match: an upstream listing that comes back empty
  # (find missing, an unreadable clone) reported "current" until 2026-09-27.
  mkdir -p "$t/empty-up"
  if check_pin "$t/vend" "$t/empty-up" >/dev/null 2>&1; then
    echo "selftest FAIL: an empty upstream tree read as a match" >&2; exit 1
  fi

  # Clone mode: the shipped workflow points at a repository URL, so the default
  # invocation must reach tree mode through a shallow clone (before 2026-09-02
  # the bare call compared SKILL.md only).
  mkdir -p "$t/repo/skills/atelier/references"
  cp "$t/up/SKILL.md" "$t/repo/skills/atelier/SKILL.md"
  cp "$t/up/references/x.md" "$t/repo/skills/atelier/references/x.md"
  ( cd "$t/repo" && git init -q . && git -c user.email=t@e.st -c user.name=t add -A \
    && git -c user.email=t@e.st -c user.name=t commit -qm 'chore: upstream' )
  if ! UPSTREAM="file://$t/repo" run "$t/vend" >/dev/null 2>&1; then
    echo "selftest FAIL: a current tree was rejected against a cloned upstream" >&2; exit 1
  fi
  rm -rf "$WORK/clone"
  # Git for Windows sets core.autocrlf=true system-wide, so a plain clone checked upstream
  # out with CRLF and a current LF copy read as behind in every file (a consumer on
  # Windows, 2026-09-28). The config is planted where that machine carries it.
  printf '[core]\n\tautocrlf = true\n' > "$t/autocrlf.gitconfig"
  if ! ( export GIT_CONFIG_SYSTEM="$t/autocrlf.gitconfig" GIT_CONFIG_GLOBAL="$t/autocrlf.gitconfig"
         UPSTREAM="file://$t/repo" run "$t/vend" ) >/dev/null 2>&1; then
    echo "selftest FAIL: a current tree was rejected against an upstream cloned under core.autocrlf=true" >&2; exit 1
  fi
  rm -rf "$WORK/clone"
  printf 'detail v1\n' > "$t/vend/references/x.md"
  if UPSTREAM="file://$t/repo" run "$t/vend" >/dev/null 2>&1; then
    echo "selftest FAIL: a stale reference passed against a cloned upstream" >&2; exit 1
  fi
  echo "selftest OK: gate rejects a stale vendored copy (file, tree, and cloned upstream), accepts a current one (also cloned under core.autocrlf=true), degrades when it cannot check"
}

if [ "${1:-}" = "--selftest" ]; then
  selftest
  exit 0
fi

run "${1:-.claude/skills/atelier}"
