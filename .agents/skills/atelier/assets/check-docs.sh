#!/usr/bin/env bash
#
# docs-check (canon 12.1): run the README's documented Verify commands so a stale
# README fails CI instead of rotting until "works on my machine" is the only real
# onboarding path.
#
# It reads the fenced ```bash blocks under the "## Verify" heading of the README and
# runs each line, but only a line that names one of the repo's own entry points:
#
#   bun run <script> [args]           a script package.json defines
#   bun test [args]                   the repo's test suite
#   bash scripts/<file> [args]        a file that exists under scripts/
#   ./scripts/<file> [args]           the same, executed directly
#   ./mvnw <phase or check> [flags]   the committed Maven wrapper: lifecycle phases,
#                                     spotless:check, pmd:check, -q/-B/-e/-ntp/-o only
#   test -f|-d|-e <path>              a path assertion
#
# Every word must be plain (letters, digits and . _ / : = @ % + , -), so a pipe, a
# redirect, ; & $ ` quotes, globs or brackets refuse the line, as does any other
# command. Every line is checked before any line runs, and each accepted line runs as
# an argument list, never through `bash -c` or eval. The README can name what the repo
# already does; it can never carry code of its own into CI. A health curl or a longer
# smoke belongs in a script under scripts/ or in package.json, which the Verify line
# then calls.
#
# Until 2026-09-26 this ran the whole block through `bash -c`, which made README text a
# second code channel with the CI runner's token; the skills.sh audits (Socket, Gen)
# flagged exactly that, and they were right.
#
# Wire it as a docs-check CI job with read-only permissions (references/governance.md);
# prove it can fail by pointing it at a README whose documented command no longer works.
#
#   bash scripts/check-docs.sh [README.md]
#
set -euo pipefail

README="${1:-README.md}"
if [ ! -f "$README" ]; then
  echo "docs-check: $README not found" >&2
  exit 1
fi

# The lines inside ```bash fences under the "## Verify" heading; a trailing comment,
# a comment line and a blank line are dropped.
block=$(awk '
  /^## / { in_verify = ($0 ~ /^##[[:space:]]+Verify([[:space:]]|$)/) }
  in_verify && /^```bash[[:space:]]*$/ { in_fence = 1; next }
  in_verify && /^```/ { in_fence = 0; next }
  in_verify && in_fence {
    sub(/[[:space:]]+#.*$/, "")
    if ($0 ~ /^[[:space:]]*(#|$)/) next
    print
  }
' "$README")

if [ -z "$block" ]; then
  echo "docs-check: no '## Verify' bash block in $README, nothing to run" >&2
  exit 0
fi

PLAIN_WORD='^[A-Za-z0-9._/:=@%+,-]+$'
scripts_loaded=0
scripts_list=""
reason=""

# The script names package.json defines, one per line (bun where it exists, else python3).
package_scripts() {
  [ -f package.json ] || return 0
  if command -v bun >/dev/null 2>&1; then
    bun -e 'const p = await Bun.file("package.json").json(); for (const k of Object.keys(p.scripts ?? {})) console.log(k)'
  elif command -v python3 >/dev/null 2>&1; then
    python3 -c 'import json; [print(k) for k in json.load(open("package.json")).get("scripts", {})]'
  else
    echo "docs-check: reading package.json scripts needs bun or python3" >&2
    return 1
  fi
}

# A path under scripts/ that names an existing file and never climbs out.
repo_script() {
  case "$1" in
    scripts/*) ;;
    *) return 1 ;;
  esac
  case "/$1/" in
    */../*) return 1 ;;
  esac
  [ -f "$1" ]
}

# Succeeds when the line is an entry point; otherwise sets `reason` and fails.
entry_point() {
  local word
  local -a w
  read -r -a w <<< "$1"
  for word in "${w[@]}"; do
    if ! [[ $word =~ $PLAIN_WORD ]]; then
      reason="'$word' is not a plain word (a pipe, redirect, quote, variable or glob is never run)"
      return 1
    fi
  done
  case "${w[0]}" in
    bun)
      case "${w[1]:-}" in
        test) return 0 ;;
        run)
          if [ "$scripts_loaded" -eq 0 ]; then
            scripts_list=$(package_scripts) || { reason="package.json scripts could not be read"; return 1; }
            scripts_loaded=1
          fi
          if [ -n "${w[2]:-}" ] && printf '%s\n' "$scripts_list" | grep -qxF -- "${w[2]}"; then return 0; fi
          reason="package.json defines no script '${w[2]:-}'"
          return 1
          ;;
      esac
      ;;
    bash)
      if repo_script "${w[1]:-}"; then return 0; fi
      reason="'${w[1]:-}' is not a file under scripts/"
      return 1
      ;;
    ./scripts/*)
      if repo_script "${w[0]#./}"; then return 0; fi
      reason="'${w[0]}' is not a file under scripts/"
      return 1
      ;;
    ./mvnw)
      if [ ! -f mvnw ]; then
        reason="there is no ./mvnw in this repo"
        return 1
      fi
      # Lifecycle phases, the two check goals and the quiet/batch flags only. A -D
      # property or a qualified plugin goal (exec:exec -Dexec.executable=...) runs any
      # executable, which made ./mvnw a way around the whole allow-list until 2026-09-27.
      for word in "${w[@]:1}"; do
        case "$word" in
          validate | compile | test | package | verify | install | clean | spotless:check | pmd:check) ;;
          -q | --quiet | -B | --batch-mode | -e | --errors | -ntp | --no-transfer-progress | -o | --offline) ;;
          *)
            reason="'./mvnw $word' is not a lifecycle phase, spotless:check, pmd:check or a quiet/batch flag (a -D property or a plugin goal can run any executable)"
            return 1
            ;;
        esac
      done
      return 0
      ;;
    test)
      case "${w[1]:-}" in
        -f | -d | -e)
          if [ "${#w[@]}" -eq 3 ]; then return 0; fi
          ;;
      esac
      reason="a path assertion is exactly 'test -f|-d|-e <path>'"
      return 1
      ;;
  esac
  reason="'${w[0]}' is not one of the repo's entry points (bun run, bun test, scripts/, ./mvnw, test)"
  return 1
}

lines=()
while IFS= read -r line; do lines+=("$line"); done <<< "$block"

refused=0
n=0
for line in "${lines[@]}"; do
  n=$((n + 1))
  if ! entry_point "$line"; then
    echo "docs-check: Verify line $n refused, $reason: $line" >&2
    refused=1
  fi
done
if [ "$refused" -ne 0 ]; then
  echo "  fix: move the command into a script under scripts/ or package.json and call it from the Verify block; nothing ran" >&2
  exit 1
fi

echo "docs-check: running the README Verify commands..."
for line in "${lines[@]}"; do
  read -r -a argv <<< "$line"
  echo "  \$ $line"
  "${argv[@]}"
done
echo "docs-check: README Verify commands passed"
