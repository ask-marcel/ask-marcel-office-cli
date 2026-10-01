#!/usr/bin/env bash
#
# Block commits if package.json declares any version as "latest" or "*".
#
# Why: "latest" / "*" are non-deterministic — `bun install` on different
# days produces different node_modules trees. The lockfile only partially
# helps, and the literal string semantically signals "always upgrade",
# which is a silent-break footgun.
#
# Add new packages with `bun add <pkg>` (runtime) or `bun add -d <pkg>`
# (dev). Bun resolves the actual latest at install time and pins it as
# `^X.Y.Z`. To bump everything to current latest deliberately, run
# `bun update` and commit the lockfile change in the same commit.
#
# Checks the root package.json and the one in every top-level folder: in a
# workspace the root manifest is plumbing and the real dependencies live in the
# folders, so checking the root alone would pass while checking nothing.
#
#   --selftest  proves a forbidden version in the root's or a folder's manifest
#               is caught and a clean tree passes (rule 15.10)
#
# See skills/atelier/references/workflow.md (Dependency hygiene) and
# SKILL.md hard rule 19.

set -euo pipefail

# Match a VALUE position (after the colon) equal to the bare strings
# "latest", "*", or a bare dist-tag ("beta", "alpha", "next", "canary",
# "rc") — all non-deterministic in exactly the way rule 19 bans.
# Anchoring on the colon keeps package NAMES out of scope (the dependency
# "next" is fine; the version "next" is not).
# Catches:  "any-pkg": "latest",   "x": "*",   "plugin": "beta"
# Permits:  "x": "^1.2.3" / "~1.2.3" / ">=1.0.0" / "^4.0.0-beta.0",  "next": "16.1.1"
#
# grep exits 1 when a manifest is clean and 2 when it cannot read it; only the
# second is an error. -H names the manifest on every line it reports.
check_manifests() {
  local f hits status violations=""
  for f in package.json */package.json; do
    if [ ! -f "$f" ]; then continue; fi
    status=0
    hits=$(grep -HnE ':[[:space:]]*"(\*|latest|beta|alpha|next|canary|rc)"' "$f") || status=$?
    if [ "$status" -gt 1 ]; then
      echo "check-package-json: cannot read $f" >&2
      return 2
    fi
    if [ -n "$hits" ]; then violations="${violations}${hits}"$'\n'; fi
  done

  if [ -z "$violations" ]; then
    return 0
  fi

  cat <<EOF >&2
  ╳ a package.json contains a forbidden version string ("latest", "*", or a bare dist-tag):

$(echo "$violations" | sed 's/^/      /')

  Atelier rule 19: every dependency declares a concrete version or range.
  Fix:
    - Replace each "latest" / "*" / bare dist-tag with the actual installed
      version (a pre-release pin like "^4.0.0-beta.0" is fine; bare "beta" is not).
    - For new packages, use \`bun add <pkg>\` (or \`bun add -d <pkg>\`)
      instead of hand-editing — Bun pins to ^X.Y.Z automatically.
    - To bump everything to current latest, run \`bun update\` and commit
      the lockfile change in the same commit.

  Bypass (rare): git commit --no-verify, with justification in commit body.
EOF
  return 1
}

selftest() {
  local tmp status=0
  tmp=$(mktemp -d "${TMPDIR:-/tmp}/check-package-json.XXXXXX")
  mkdir -p "$tmp/clean/my pkg" "$tmp/dirty-folder/my pkg" "$tmp/dirty-root/pkg"
  printf '{ "dependencies": { "next": "16.1.1" } }\n' > "$tmp/clean/package.json"
  printf '{ "dependencies": { "a": "^1.2.3" } }\n' > "$tmp/clean/my pkg/package.json"
  printf '{ "private": true }\n' > "$tmp/dirty-folder/package.json"
  printf '{ "dependencies": { "a": "latest" } }\n' > "$tmp/dirty-folder/my pkg/package.json"
  printf '{ "dependencies": { "a": "*" } }\n' > "$tmp/dirty-root/package.json"
  printf '{ "dependencies": { "a": "^1.2.3" } }\n' > "$tmp/dirty-root/pkg/package.json"
  if ! (cd "$tmp/clean" && check_manifests); then
    echo "check-package-json: SELFTEST FAILED, a clean tree was rejected" >&2
    status=1
  fi
  if (cd "$tmp/dirty-folder" && check_manifests 2>/dev/null); then
    echo "check-package-json: SELFTEST FAILED, \"latest\" in a folder's package.json was accepted" >&2
    status=1
  fi
  if (cd "$tmp/dirty-root" && check_manifests 2>/dev/null); then
    echo "check-package-json: SELFTEST FAILED, \"*\" in the root package.json was accepted" >&2
    status=1
  fi
  rm -rf "$tmp"
  if [ "$status" -eq 0 ]; then
    echo "check-package-json: selftest passed (a forbidden version in the root's or a folder's manifest is rejected)."
  fi
  return "$status"
}

if [ "${1:-}" = "--selftest" ]; then
  selftest
  exit $?
fi

check_manifests
