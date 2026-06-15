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
# See skills/atelier/references/workflow.md (Dependency hygiene) and
# SKILL.md hard rule 19.

set -euo pipefail

if [ ! -f package.json ]; then
  exit 0
fi

# Match a value position equal to the bare strings "latest" or "*".
# Catches:  "any-pkg": "latest",   "x": "*"
# Permits:  "x": "^1.2.3" / "1.2.3" / "~1.2.3" / ">=1.0.0" etc.
violations=$(grep -nE '"\*"|"latest"' package.json || true)

if [ -z "$violations" ]; then
  exit 0
fi

cat <<EOF >&2
  ╳ package.json contains forbidden version "latest" or "*":

$(echo "$violations" | sed 's/^/      /')

  Atelier rule 19: every dependency declares a concrete version or range.
  Fix:
    - Replace each "latest" / "*" with the actual installed version.
    - For new packages, use \`bun add <pkg>\` (or \`bun add -d <pkg>\`)
      instead of hand-editing — Bun pins to ^X.Y.Z automatically.
    - To bump everything to current latest, run \`bun update\` and commit
      the lockfile change in the same commit.

  Bypass (rare): git commit --no-verify, with justification in commit body.
EOF
exit 1
