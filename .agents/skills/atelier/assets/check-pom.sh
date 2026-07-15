#!/usr/bin/env bash
#
# Block Maven version drift in every tracked pom.xml (Java variant, rule 19):
#
#   1. Version ranges, e.g. <version>[1.0,)</version> or (,2.0]:
#      Maven resolves them to whatever is newest that day, which is the
#      "latest" footgun with different syntax. Builds must be reproducible.
#   2. -SNAPSHOT versions in <parent>, <dependencies>, or <plugins> blocks:
#      a snapshot is mutable upstream, so the same commit builds differently
#      over time. The project's own <version> may be a SNAPSHOT during
#      development; third-party coordinates may not.
#
# The maven-enforcer-plugin (requireUpperBoundDeps, banSnapshots on release)
# is the build-time authority; this hook is the fast pre-commit echo of it.
# See skills/atelier/references/java-quarkus.md (pom.xml conventions).

set -euo pipefail

status=0

while IFS= read -r pom; do
  [ -f "$pom" ] || continue

  # 1. Any bracket or parenthesis inside a <version> element is a range.
  if grep -nE '<version>[^<]*[][()]' "$pom"; then
    echo "  ╳ $pom declares a version range; pin an exact version (rule 19)" >&2
    status=1
  fi

  # 2. A -SNAPSHOT <version> inside parent/dependencies/plugins blocks (never
  #    third-party). Matching only <version> elements keeps prose such as an
  #    enforcer <message> mentioning -SNAPSHOT from tripping the gate.
  snapshot_hits=$(awk '
    /<parent>|<dependencies>|<plugins>/       { depth += 1 }
    depth > 0 && /<version>[^<]*-SNAPSHOT/    { printf "    line %d: %s\n", NR, $0 }
    /<\/parent>|<\/dependencies>|<\/plugins>/ { if (depth > 0) depth -= 1 }
  ' "$pom")
  if [ -n "$snapshot_hits" ]; then
    echo "  ╳ $pom pins a -SNAPSHOT dependency; use a released version (rule 19)" >&2
    echo "$snapshot_hits" >&2
    status=1
  fi
# --others --exclude-standard: a brand-new pom is checked before its first
# commit too, not only once tracked.
done < <(git ls-files --cached --others --exclude-standard '*pom.xml' 'pom.xml' | sort -u)

exit "$status"
