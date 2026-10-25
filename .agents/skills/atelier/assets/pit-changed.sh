#!/usr/bin/env bash
#
# Run PIT on the classes that changed under domain/ and usecases/ (plus any
# uncommitted edits and untracked files): the mutation step of the Java CI
# gate (assets/ci-java.yml, every pull request and push). The full sweep is
# never a commit gate; assets/mutation-java.yml runs it once a day. The Java
# twin of mutate-changed.sh, same base resolution, same loud failures.
#
# The base ref: BASE= wins; under GitHub Actions the range is the pull
# request's base branch or, on a push, github.event.before..HEAD (ci-java.yml
# exports it as GITHUB_EVENT_BEFORE; the zero SHA of a new branch and an
# unresolvable SHA fall back to HEAD~1). Locally the default is the branch's
# push ref (`@{push}`, the classes changed since your last push; CI's run over
# the pull request is the gate), else origin/main for a branch never pushed.
# Skip the ref refresh with MUTATE_NO_FETCH=1.
#
# The canonical pom reads `<targetClasses>${pitest.targetClasses}</targetClasses>`
# with the two package globs as the property default, so this script narrows
# the run by overriding the property with the changed classes' names.
#
# See skills/atelier/references/java-quarkus.md (Testing; Gates and hooks).

set -euo pipefail

zero_sha=0000000000000000000000000000000000000000
if [ -n "${BASE:-}" ]; then
  :
elif [ -n "${GITHUB_BASE_REF:-}" ]; then
  BASE="origin/${GITHUB_BASE_REF}"
elif [ "${GITHUB_EVENT_NAME:-}" = "push" ]; then
  before="${GITHUB_EVENT_BEFORE:-}"
  if [ -n "$before" ] && [ "$before" != "$zero_sha" ] && git rev-parse --quiet --verify "${before}^{commit}" >/dev/null 2>&1; then
    BASE="$before"
  else
    BASE=HEAD~1
  fi
elif push_ref=$(git rev-parse --abbrev-ref --symbolic-full-name '@{push}' 2>/dev/null) \
     && git rev-parse --quiet --verify "${push_ref}^{commit}" >/dev/null 2>&1; then
  BASE="$push_ref"
else
  BASE=origin/main
fi

if [ -z "${MUTATE_NO_FETCH:-}" ] && [ "${BASE#origin/}" != "$BASE" ]; then
  git fetch --quiet origin "${BASE#origin/}" || true
fi

# An unknown BASE would make every diff below fail and the `|| true` on the
# pipeline would turn that into "no classes changed" plus exit 0. Fail loudly.
if ! git rev-parse --verify --quiet "$BASE^{commit}" >/dev/null; then
  echo "pit-changed: base ref '$BASE' does not exist (fetch it, or set BASE=)" >&2
  exit 1
fi

echo "pit-changed: base ${BASE} $(git rev-parse --short "$BASE")" \
     "($(git log -1 --format=%cr "$BASE")), HEAD +$(git rev-list --count "$BASE"..HEAD)"

# Changed, uncommitted, staged, and untracked sources in the mutation scope
# (domain and usecases packages, ports excluded: interfaces carry no logic).
files=$( {
  git diff --name-only --diff-filter=ACMR "$BASE"...HEAD
  git diff --name-only --diff-filter=ACMR HEAD
  git diff --cached --name-only --diff-filter=ACMR
  git ls-files --others --exclude-standard
} | sort -u \
  | grep -E '(^|/)src/main/java/.*/(domain|usecases)/.*\.java$' \
  | grep -vE '/ports/' \
  || true)

# A module's sources (svc/src/main/java/...) need PIT run in that module; the root
# run would target classes it cannot see. Refuse loudly instead of the vacuous
# "nothing changed" this script printed for every multi-module repo until 2026-09-27.
nested=$(echo "$files" | grep -vE '^src/main/java/' | grep -v '^$' || true)
if [ -n "$nested" ]; then
  echo "pit-changed: changed classes live in a module, not the root build:" >&2
  echo "$nested" | sed 's/^/    /' >&2
  echo "  scope PIT per module (references/java-quarkus.md): run this script from each module's directory with its own ./mvnw -pl, or split the CI step per module" >&2
  exit 1
fi

if [ -z "$files" ]; then
  echo "pit-changed: no classes in mutation scope changed since ${BASE}"
  exit 0
fi

# Path to fully qualified class name: src/main/java/com/x/domain/Money.java -> com.x.domain.Money,
# plus its nested classes (com.x.domain.Money$*): logic in a nested class was never
# mutated until 2026-09-27, so a weak test scored 100 on the outer class alone.
classes=$(echo "$files" | sed -e 's|^src/main/java/||' -e 's|\.java$||' -e 's|/|.|g' | sed 's|.*|&,&$*|' | paste -sd, -)
count=$(echo "$files" | wc -l | tr -d ' ')
echo "pit-changed: targeting ${count} class(es): ${classes}"

./mvnw -q test-compile org.pitest:pitest-maven:mutationCoverage -Dpitest.targetClasses="${classes}"

# PIT's mutationThreshold judges the whole run's score, so a weak class passes beside strong ones
# (2026-10-23). The per-class gate holds each targeted class to the same threshold.
python3 scripts/check-pit-score.py
