#!/usr/bin/env bash
#
# Block Maven version drift in every tracked pom.xml (Java variant, rule 19):
#
#   1. Version ranges, e.g. <version>[1.0,)</version> or (,2.0]:
#      Maven resolves them to whatever is newest that day, which is the
#      "latest" footgun with different syntax. Builds must be reproducible.
#      Read in <version> elements AND in version properties (<junit.version>),
#      where the canonical pom keeps every pin: a range in a property passed
#      until 2026-09-27.
#   2. -SNAPSHOT versions in <parent>, <dependencies>, or <plugins> blocks, or
#      in a version property: a snapshot is mutable upstream, so the same
#      commit builds differently over time. A single-module project's own
#      <version> may be a SNAPSHOT during development; in a multi-module repo a
#      child's <parent> is that SNAPSHOT and both this gate and the enforcer
#      reject it, so keep release versions there and bump them at release.
#   3. A mock library (rule 13): Mockito, EasyMock, PowerMock, JMockit, JMock
#      or the Quarkus mockito (both names) and panache-mock extensions, as a
#      <groupId> or an <artifactId>. Hand-written fakes implement the ports. The
#      enforcer's bannedDependencies is the build-time authority (it sees
#      transitives too); this is the fast echo that stops the declaration
#      before the commit.
#
# The maven-enforcer-plugin (requireReleaseDeps, banDynamicVersions,
# bannedDependencies) is the build-time authority; this hook is the fast
# pre-commit echo of it.
# See skills/atelier/references/java-quarkus.md (pom.xml conventions).

set -euo pipefail

status=0

while IFS= read -r pom; do
  [ -f "$pom" ] || continue

  # 1. Any bracket or parenthesis inside a <version> element or a version
  #    property (<x.version>, <x-version>) is a range.
  if grep -nE '<([A-Za-z0-9_.-]*[.-])?version>[^<]*[][()]' "$pom"; then
    echo "  ╳ $pom declares a version range; pin an exact version (rule 19)" >&2
    status=1
  fi

  # 2. A -SNAPSHOT <version> inside parent/dependencies/plugins blocks (never
  #    third-party), or a -SNAPSHOT version property. Matching only version
  #    elements keeps prose such as an enforcer <message> mentioning -SNAPSHOT
  #    from tripping the gate.
  snapshot_hits=$(awk '
    /<parent>|<dependencies>|<plugins>/       { depth += 1 }
    depth > 0 && /<version>[^<]*-SNAPSHOT/    { printf "    line %d: %s\n", NR, $0 }
    /<[A-Za-z0-9_.-]*[.-]version>[^<]*-SNAPSHOT/ { printf "    line %d: %s\n", NR, $0 }
    /<\/parent>|<\/dependencies>|<\/plugins>/ { if (depth > 0) depth -= 1 }
  ' "$pom")
  if [ -n "$snapshot_hits" ]; then
    echo "  ╳ $pom pins a -SNAPSHOT dependency; use a released version (rule 19)" >&2
    echo "$snapshot_hits" >&2
    status=1
  fi

  # 3. A mock library declared anywhere in the pom (rule 13). <exclude> entries
  #    of the enforcer's own ban list are not declarations and do not match.
  if grep -nE '<groupId>(org\.mockito|org\.easymock|org\.powermock|org\.jmockit|org\.jmock)</groupId>|<artifactId>(mockito-[a-z0-9-]+|easymock|easymockclassextension|powermock-[a-z0-9-]+|jmockit|jmock(-[a-z0-9-]+)?|quarkus-junit5?-mockito|quarkus-panache-mock)</artifactId>' "$pom"; then
    echo "  ╳ $pom declares a mock library; hand-written fakes implement the ports (rule 13)" >&2
    status=1
  fi
# --others --exclude-standard: a brand-new pom is checked before its first
# commit too, not only once tracked.
done < <(git ls-files --cached --others --exclude-standard '*pom.xml' 'pom.xml' | sort -u)

exit "$status"
