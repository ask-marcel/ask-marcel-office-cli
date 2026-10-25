#!/usr/bin/env bash
#
# Rule 26 tripwire: identity lives in commit metadata, never in file contents.
#
# Checks the STAGED ADDED LINES (like the staged secret scan), so it blocks
# a mention entering history without flooding a brownfield tree. `--all` scans
# every tracked text file instead (CI, and the adopt-mode audit).
#
# What it looks for, case-insensitively, as whole words:
#   - the committer's full name, in both orders (git config user.name)
#   - the committer's email (git config user.email)
#   - under --all, every author and committer name and email in the history
#   - every entry of IDENTITY_DENYLIST (newline- or comma-separated: employer
#     and client names). An environment variable, never a tracked file, since a
#     tracked denylist would itself name what the rule forbids.
# A one-word git name is a handle, which rule 26 permits, and is skipped; so is
# a GitHub noreply address (<handle>@users.noreply.github.com). CODEOWNERS and
# .mailmap are exempt: rule 26 names them as the two places a person belongs.
# A lone first name stays a review duty: a one-word match would flag "will"
# and "grace" in prose.
#
# This is a tripwire, not a proof. In a public repository a hit is a hard stop; the
# fix is a neutral handle, never an inline suppression (rule 15).
#
# Public blocks, private warns (2026-10-18). A name in a public repository's files is
# public the moment it is pushed; in a private one it is a debt to clear before the
# repository goes public. So a hit exits 1 in a public repository and in a private
# one prints the same hits as a warning and exits 0. The visibility is CI's
# REPO_VISIBILITY when set (the host's own answer, `private` or `public`, which every
# shipped CI workflow passes from github.event.repository.private), else the committed
# declaration .github/visibility, one line reading `private`. Neither, or any other
# word, is public. A file saying private on a public host exits 1 with or without a
# hit, so a repository made public cannot keep warning.
#
#   bash scripts/check-identity.sh          # staged added lines (the hook)
#   bash scripts/check-identity.sh --all    # every tracked text file (CI, adopt audit)

set -euo pipefail

MODE="${1:-staged}"

declaration="$(git rev-parse --show-toplevel 2>/dev/null || pwd)/.github/visibility"
declared=""
if [ -f "$declaration" ]; then
  declared="$(head -n 1 "$declaration" | tr -d '[:space:]' | tr '[:upper:]' '[:lower:]')"
fi
host="$(printf '%s' "${REPO_VISIBILITY:-}" | tr -d '[:space:]' | tr '[:upper:]' '[:lower:]')"
if [ -n "$host" ] && [ "$host" != private ] && [ "$declared" = private ]; then
  echo "check-identity: .github/visibility says private, but the host says this repository is $host (hard rule 26)" >&2
  echo "  fix: delete .github/visibility; a public repository blocks every name (references/workflow.md, Commit identity)" >&2
  exit 1
fi
visibility="${host:-$declared}"
[ "$visibility" = private ] || visibility=public

# The identity set: one string per line, trimmed, at least four characters.
# Git names count only with a space in them (a one-word name is a handle);
# emails count unless they are GitHub noreply handles; denylist entries always.
identities() {
  {
    { git config user.name 2>/dev/null || true; } | awk '/ /'
    git config user.email 2>/dev/null || true
    if [ "$MODE" = "--all" ]; then
      { git log --format='%an%n%cn' 2>/dev/null || true; } | awk '/ /'
      git log --format='%ae%n%ce' 2>/dev/null || true
    fi
    printf '%s\n' "${IDENTITY_DENYLIST:-}" | tr ',' '\n'
  } | sed 's/^[[:space:]]*//; s/[[:space:]]*$//' \
    | grep -v -i '@users\.noreply\.github\.com$' \
    | awk 'length($0) >= 4' | sort -u
}

# A multi-word name in both orders: "First Last" and "Last First".
patterns() {
  identities | while IFS= read -r id; do
    printf '%s\n' "$id"
    case "$id" in
      *" "*) printf '%s %s\n' "${id#* }" "${id%% *}" ;;
    esac
  done | sort -u
}

# One stream of `path: content` (staged) or `path:N:content` (--all) lines,
# CODEOWNERS and .mailmap dropped, binaries skipped.
lines() {
  if [ "$MODE" = "--all" ]; then
    git ls-files -z | xargs -0 grep -I -H -n '' 2>/dev/null \
      | awk -F: '$1 !~ /(^|\/)(CODEOWNERS|\.mailmap)$/' || true
  else
    git diff --cached -U0 --diff-filter=ACMR \
      | awk '/^diff --git /{h=1; next} h && /^\+\+\+ /{f=substr($0,7); next} /^@@/{h=0; next} !h && /^\+/{print f": "substr($0,2)}' \
      | awk -F': ' '$1 !~ /(^|\/)(CODEOWNERS|\.mailmap)$/' || true
  fi
}

tmp="$(mktemp)"
trap 'rm -f "$tmp"' EXIT
patterns > "$tmp"
if [ ! -s "$tmp" ]; then
  echo "check-identity: no identity to look for (git config user.name/user.email unset and IDENTITY_DENYLIST empty)" >&2
  exit 0
fi

hits="$(lines | grep -F -i -w -f "$tmp" || true)"
if [ -n "$hits" ] && [ "$visibility" = private ]; then
  echo "check-identity: warning: a person, employer, or client is named in file contents (hard rule 26 blocks this in a public repository; this one is private):" >&2
  printf '%s\n' "$hits" | sed 's/^/  /' >&2
  echo "  before the repository goes public: a neutral handle for each, then scrub history (references/workflow.md, Commit identity)" >&2
  if [ "${GITHUB_ACTIONS:-}" = true ]; then
    echo "::warning title=rule 26::a person, employer, or client is named in file contents; this blocks once the repository is public"
  fi
  exit 0
fi
if [ -n "$hits" ]; then
  echo "check-identity: a person, employer, or client is named in file contents (hard rule 26):" >&2
  printf '%s\n' "$hits" | sed 's/^/  /' >&2
  echo "  fix: a neutral handle in file contents; CODEOWNERS and .mailmap are where a person belongs (references/workflow.md, Commit identity)" >&2
  exit 1
fi
echo "check-identity: ok ($MODE)"
