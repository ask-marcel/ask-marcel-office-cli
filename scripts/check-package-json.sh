#!/usr/bin/env bash
#
# Gate 2 of the fast hook, run again in CI: the manifests and the toolchain.
#
#   1. Rule 19: no package.json declares a version as "latest", "*", "x", a
#      bare dist-tag, an npm: alias with no version of its own, or (in a
#      dependency block) the empty string.
#   2. Rule 5: no lockfile of another package manager is tracked or staged
#      anywhere (package-lock.json, npm-shrinkwrap.json, yarn.lock,
#      pnpm-lock.yaml). bun.lock is the lockfile.
#   3. Rule 5: no "scripts" entry calls node, npm, npx, pnpm, yarn or vite
#      directly, as the command or as a segment after &&, ||, ; or |. An env
#      prefix (LINT_STRICT=1 ...) is allowed. `bunx vite` and `bun run node-x`
#      are not direct calls and pass.
#
# Checks 2 and 3 come from the atelier asset check-package-json.sh, merged by
# hand: this copy keeps its own manifest list and its --selftest.
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
#   --selftest  proves that each check rejects its violation and that a clean
#               tree passes (rule 15.10)
#
# See skills/atelier/references/workflow.md (Dependency hygiene) and
# SKILL.md hard rules 5 and 19.

set -euo pipefail

# Match a VALUE position (after the colon) equal to the bare strings
# "latest", "*", "x" / "X", a bare dist-tag ("beta", "alpha", "next",
# "canary", "rc"), or an npm: alias that resolves to one or names no version
# at all. All of them install whatever is newest, in exactly the way rule 19
# bans. Anchoring on the colon keeps package NAMES out of scope (the
# dependency "next" is fine; the version "next" is not).
# Catches:  "any-pkg": "latest",   "x": "*",   "z": "x",   "plugin": "beta",
#           "a": "npm:b@latest",   "a": "npm:b",   "a": "npm:@s/b"
# Permits:  "x": "^1.2.3" / "~1.2.3" / ">=1.0.0" / "^4.0.0-beta.0",  "next": "16.1.1",
#           "a": "npm:b@^1.2.3",   "w": "workspace:*"
VERSION_RE=':[[:space:]]*"(\*|[xX]|latest|beta|alpha|next|canary|rc|npm:(@[^/"]+/)?[^@"]+(@(latest|\*|[xX])?)?)"'

# The empty string is a forbidden version too, but only in a block of
# versions: "description": "" is not a finding. awk reads the dependency
# blocks (and overrides and resolutions, which hold versions too) and prints
# each "" value in them, as file:line:text.
empty_versions() {
  awk '
    FNR == 1 { inblock = 0 }
    {
      s = $0
      while (match(s, /"(dependencies|devDependencies|peerDependencies|optionalDependencies|overrides|resolutions)"[[:space:]]*:[[:space:]]*\{/)) {
        rest = substr(s, RSTART + RLENGTH)
        close_at = index(rest, "}")
        if (close_at > 0) {
          if (substr(rest, 1, close_at) ~ /:[[:space:]]*""/) print FILENAME ":" FNR ":" $0
          s = substr(rest, close_at + 1); continue
        }
        inblock = 1; s = ""
      }
      if (!inblock) next
      if ($0 ~ /^[[:space:]]*\}/) { inblock = 0; next }
      if ($0 ~ /:[[:space:]]*""/) print FILENAME ":" FNR ":" $0
    }' "$@"
}

# grep exits 1 when a manifest is clean and 2 when it cannot read it; only the
# second is an error. -H names the manifest on every line it reports.
check_manifests() {
  local f hits status violations=""
  for f in package.json */package.json; do
    if [ ! -f "$f" ]; then continue; fi
    status=0
    hits=$(grep -HnE "$VERSION_RE" "$f") || status=$?
    if [ "$status" -gt 1 ]; then
      echo "check-package-json: cannot read $f" >&2
      return 2
    fi
    if [ -n "$hits" ]; then violations="${violations}${hits}"$'\n'; fi
    hits=$(empty_versions "$f")
    if [ -n "$hits" ]; then violations="${violations}${hits}"$'\n'; fi
  done

  if [ -z "$violations" ]; then
    return 0
  fi

  cat <<EOF >&2
  ╳ a package.json contains a forbidden version string ("latest", "*", "x", "", a bare dist-tag, or an npm: alias with no version of its own):

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

# Rule 5: a lockfile of another package manager, tracked or not yet tracked,
# anywhere in the repo, means that an install ran through npm, yarn or pnpm.
# Installed packages ship their own lockfiles, so node_modules is out of scope.
check_lockfiles() {
  local lockfiles
  lockfiles=$( { git ls-files --cached --others --exclude-standard 2>/dev/null || true; } \
    | grep -v -E '(^|/)node_modules/' \
    | grep -E '(^|/)(package-lock\.json|npm-shrinkwrap\.json|yarn\.lock|pnpm-lock\.yaml)$' || true)
  if [ -z "$lockfiles" ]; then
    return 0
  fi
  echo "  ╳ a lockfile of another package manager is in the repo; Bun only, bun.lock is the lockfile (rule 5):" >&2
  echo "$lockfiles" | sed 's/^/      /' >&2
  return 1
}

# Rule 5: a "scripts" entry that calls another runtime or package manager
# directly. Each script value is cut at &&, ||, ; and |, and each segment,
# after its env prefixes, must not start with one of the banned commands.
check_scripts() {
  local f manifests=() hits
  for f in package.json */package.json; do
    if [ -f "$f" ]; then manifests+=("$f"); fi
  done
  if [ "${#manifests[@]}" -eq 0 ]; then
    return 0
  fi
  hits=$(awk '
    function banned(v,   n, parts, i, seg) {
      gsub(/&&|\|\||;|\|/, "\n", v)
      n = split(v, parts, "\n")
      for (i = 1; i <= n; i++) {
        seg = parts[i]
        sub(/^[[:space:]]+/, "", seg)
        while (match(seg, /^[A-Za-z_][A-Za-z0-9_]*=[^[:space:]]*[[:space:]]+/)) seg = substr(seg, RLENGTH + 1)
        if (seg ~ /^(node|npm|npx|pnpm|yarn|vite)([[:space:]]|$)/) return 1
      }
      return 0
    }
    function check_line(line,   pair, v) {
      while (match(line, /"[^"]*"[[:space:]]*:[[:space:]]*"[^"]*"/)) {
        pair = substr(line, RSTART, RLENGTH); line = substr(line, RSTART + RLENGTH)
        v = pair; sub(/^"[^"]*"[[:space:]]*:[[:space:]]*"/, "", v); sub(/"$/, "", v)
        if (banned(v)) print FILENAME ":" FNR ": " pair
      }
    }
    FNR == 1 { inb = 0 }
    {
      if (!inb && match($0, /"scripts"[[:space:]]*:[[:space:]]*\{/)) {
        rest = substr($0, RSTART + RLENGTH)
        if (index(rest, "}") > 0) { check_line(substr(rest, 1, index(rest, "}"))); next }
        inb = 1; next
      }
      if (!inb) next
      if ($0 ~ /^[[:space:]]*\}/) { inb = 0; next }
      check_line($0)
    }' "${manifests[@]}")
  if [ -z "$hits" ]; then
    return 0
  fi
  echo "  ╳ a package.json script calls node, npm, npx, pnpm, yarn or vite directly; bun run, bunx and bun test are the toolchain (rule 5):" >&2
  echo "$hits" | sed 's/^/      /' >&2
  return 1
}

check_all() {
  local status=0
  check_lockfiles || status=1
  check_scripts || status=1
  check_manifests || status=1
  return "$status"
}

# A case of the selftest: a tree whose manifests the caller wrote must pass
# (expect 0) or fail (expect 1) all checks together.
selftest_case() {
  local dir=$1 expect=$2 what=$3 status=0
  (cd "$dir" && check_all 2>/dev/null) || status=1
  if [ "$status" -ne "$expect" ]; then
    if [ "$expect" -eq 0 ]; then
      echo "check-package-json: SELFTEST FAILED, $what was rejected" >&2
    else
      echo "check-package-json: SELFTEST FAILED, $what was accepted" >&2
    fi
    return 1
  fi
  return 0
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
  # The clean tree also holds what the new checks must let pass: an empty
  # value outside the dependency blocks, an npm: alias with a range, a
  # workspace range, and scripts that run vite and a node-named script
  # through bun.
  printf '{\n  "description": "",\n  "scripts": {\n    "dev": "bunx vite",\n    "x": "LINT_STRICT=1 bun run node-x && bun test"\n  },\n  "dependencies": {\n    "a": "npm:b@^1.2.3",\n    "w": "workspace:*"\n  }\n}\n' > "$tmp/clean/my pkg/package.json"
  mkdir -p "$tmp/lockfile/pkg" "$tmp/script" "$tmp/empty" "$tmp/x" "$tmp/alias"
  printf '{ "private": true }\n' > "$tmp/lockfile/package.json"
  printf '\n' > "$tmp/lockfile/pkg/yarn.lock"
  git -C "$tmp/lockfile" init -q
  printf '{\n  "scripts": {\n    "build": "LINT_STRICT=1 node build.js"\n  }\n}\n' > "$tmp/script/package.json"
  printf '{\n  "dependencies": {\n    "a": ""\n  }\n}\n' > "$tmp/empty/package.json"
  printf '{ "dependencies": { "a": "x" } }\n' > "$tmp/x/package.json"
  printf '{ "dependencies": { "a": "npm:b" } }\n' > "$tmp/alias/package.json"
  selftest_case "$tmp/clean" 0 "a clean tree" || status=1
  selftest_case "$tmp/dirty-folder" 1 "\"latest\" in a folder's package.json" || status=1
  selftest_case "$tmp/dirty-root" 1 "\"*\" in the root package.json" || status=1
  selftest_case "$tmp/lockfile" 1 "a yarn.lock in a folder" || status=1
  selftest_case "$tmp/script" 1 "a script that calls node" || status=1
  selftest_case "$tmp/empty" 1 "an empty version in a dependency block" || status=1
  selftest_case "$tmp/x" 1 "the version \"x\"" || status=1
  selftest_case "$tmp/alias" 1 "an npm: alias with no version" || status=1
  rm -rf "$tmp"
  if [ "$status" -eq 0 ]; then
    echo "check-package-json: selftest passed (a clean tree passes; \"latest\", \"*\", \"x\", \"\", a bare npm: alias, a foreign lockfile and a script that calls node are each rejected)."
  fi
  return "$status"
}

if [ "${1:-}" = "--selftest" ]; then
  selftest
  exit $?
fi

check_all
