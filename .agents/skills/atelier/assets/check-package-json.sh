#!/usr/bin/env bash
#
# Gate 2 of the fast hook, re-run in CI: the manifests and the toolchain.
#
#   1. Rule 19: no package.json declares a version as "latest", "*", a bare
#      dist-tag, or an npm: alias resolving to one.
#   2. Rule 5: no non-bun lockfile is tracked or staged anywhere
#      (package-lock.json, npm-shrinkwrap.json, yarn.lock, pnpm-lock.yaml).
#   3. Rule 5: no "scripts" entry calls node, npm, npx, pnpm, yarn or vite
#      directly, as the command or as a segment after &&, ||, ; or |, an env
#      prefix (LINT_STRICT=1 ...) allowed. `bunx vite` and `bun run node-x`
#      are not direct calls and pass; the rule says "directly".
#
# On the version strings:
# Why: "latest" / "*" are non-deterministic, `bun install` on different
# days produces different node_modules trees. The lockfile only partially
# helps, and the literal string semantically signals "always upgrade",
# which is a silent-break footgun.
#
# Add new packages with `bun add <pkg>` (runtime) or `bun add -d <pkg>`
# (dev). Bun resolves the actual latest at install time and pins it as
# `^X.Y.Z`. To bump everything to current latest deliberately, run
# `bun update` and commit the lockfile change in the same commit.
#
# See skills/atelier/references/workflow.md (Dependency hygiene) and
# SKILL.md hard rules 5 and 19.

set -euo pipefail

status=0

# Every manifest in the repo, not just the root one: in a monorepo the
# dependencies live in apps/* and packages/*, so a root-only read passes a
# repo whose workspaces pin "latest" (found in a real consumer repo,
# 2026-08-30 field test). node_modules and .git are excluded.
manifests=$(find . -name package.json -not -path '*/node_modules/*' -not -path '*/.git/*' | sort)

# 2. A lockfile of another package manager, tracked or staged, anywhere in the
#    repo: it means an install ran through npm, yarn or pnpm (rule 5).
#    Installed packages ship their own lockfiles, so node_modules is out of
#    scope whether or not the repo ignores it.
lockfiles=$( { git ls-files --cached --others --exclude-standard 2>/dev/null || true; } \
  | grep -v -E '(^|/)node_modules/' \
  | grep -E '(^|/)(package-lock\.json|npm-shrinkwrap\.json|yarn\.lock|pnpm-lock\.yaml)$' || true)
if [ -n "$lockfiles" ]; then
  echo "  ╳ a lockfile of another package manager is in the repo; Bun only, bun.lock is the lockfile (rule 5):" >&2
  echo "$lockfiles" | sed 's/^/      /' >&2
  status=1
fi

# 3. A "scripts" entry that calls another runtime or package manager directly.
if [ -n "$manifests" ]; then
  script_hits=$(echo "$manifests" | tr '\n' '\0' | xargs -0 awk '
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
    }' || true)
  if [ -n "$script_hits" ]; then
    echo "  ╳ a package.json script calls node, npm, npx, pnpm, yarn or vite directly; bun run, bunx and bun test are the toolchain (rule 5):" >&2
    echo "$script_hits" | sed 's/^/      /' >&2
    status=1
  fi
fi

if [ -z "$manifests" ]; then
  exit "$status"
fi

# Match a VALUE position (after the colon) equal to the bare strings
# "latest", "*", "x" / "X", the empty string, or a bare dist-tag ("beta",
# "alpha", "next", "canary", "rc"), or an npm: alias resolving to one or naming
# no version at all, all of which install whatever is newest, in exactly the
# way rule 19 bans ("" and "x" passed until 2026-09-27, the blank-pin trap).
# Only the four dependency blocks are read, so a version-shaped value elsewhere
# (publishConfig.tag: "next", an engines field, a script) is not a finding (a
# false positive found 2026-09-02).
# Catches:  "any-pkg": "latest",   "x": "*",   "y": "",   "z": "x",   "plugin": "beta",
#           "a": "npm:b@latest",   "a": "npm:b",   "a": "npm:@s/b"
# Permits:  "x": "^1.2.3" / "~1.2.3" / ">=1.0.0" / "^4.0.0-beta.0",  "next": "16.3.6",
#           "a": "npm:b@^1.2.3",   "w": "workspace:*"
violations=$(echo "$manifests" | tr '\n' '\0' | xargs -0 awk '
  BEGIN { V = ":[[:space:]]*\"(\\*|[xX]|latest|beta|alpha|next|canary|rc|npm:(@[^/\"]+\\/)?[^@\"]+(@(latest|\\*|[xX])?)?)?\"" }
  FNR == 1 { inblock = 0 }
  {
    s = $0
    while (match(s, /"(dependencies|devDependencies|peerDependencies|optionalDependencies)"[[:space:]]*:[[:space:]]*\{/)) {
      rest = substr(s, RSTART + RLENGTH)
      close_at = index(rest, "}")
      if (close_at > 0) {                         # a one-line block: test just its body
        if (substr(rest, 1, close_at) ~ V) print FILENAME ":" FNR ":" $0
        s = substr(rest, close_at + 1); continue
      }
      inblock = 1; s = ""                         # a multi-line block opens here
    }
    if (!inblock) next
    if ($0 ~ /^[[:space:]]*\}/) { inblock = 0; next }
    if ($0 ~ V) print FILENAME ":" FNR ":" $0
  }' || true)

if [ -z "$violations" ]; then
  exit "$status"
fi

cat <<EOF >&2
  ╳ a package.json declares a forbidden version string ("latest", "*", "x", "", a bare dist-tag, or an npm: alias with none of its own):

$(echo "$violations" | sed 's/^/      /')

  Atelier rule 19: every dependency declares a concrete version or range.
  Fix:
    - Replace each "latest" / "*" / bare dist-tag with the actual installed
      version (a pre-release pin like "^4.0.0-beta.0" is fine; bare "beta" is not).
    - For new packages, use \`bun add <pkg>\` (or \`bun add -d <pkg>\`)
      instead of hand-editing: Bun pins to ^X.Y.Z automatically.
    - To bump everything to current latest, run \`bun update\` and commit
      the lockfile change in the same commit.

  Bypass (rare): git commit --no-verify, with justification in commit body.
EOF
exit 1
