#!/usr/bin/env bash
#
# Rule 29 tripwire: every outbound call in an infra adapter has a deadline.
#
# Checks files the STAGED DIFF touches under the infra layer; `--all` scans the
# whole layer (adopt-mode audit). Per-call heuristic on the TypeScript side: a
# deadline marker must appear within the 8 lines after each call, comment lines
# stripped first, so a `// TODO timeout` cannot satisfy it and a marker on an
# unrelated call elsewhere in the file cannot vouch for this one. The Java side
# stays per-file.
#
#   TS   call markers:     fetch(  |  globalThis.fetch(   (the doctrine's idiom,
#                          references/bun-typescript.md; a member call such as
#                          deps.fetch( is a known gap, name the wrapper instead)
#        deadline markers: AbortSignal.timeout(  |  signal:
#   Java call marker:      HttpClient.new
#        deadline markers: .timeout( | connectTimeout
#
# A tripwire, not a proof: it cannot check the timeout VALUE or SDK clients it
# does not know. Test files are exempt (they call fakes). No inline suppression;
# a file that genuinely delegates its deadline should name the wrapper so a
# marker appears (see skills/atelier/references/reliability.md).

set -euo pipefail

MODE="${1:-staged}"

candidate_files() {
  if [ "$MODE" = "--all" ]; then
    { find src/infra -type f \( -name '*.ts' -o -name '*.tsx' \) 2>/dev/null;
      find src/main/java -type f -name '*.java' -path '*infra*' 2>/dev/null; } || true
  else
    git diff --cached --name-only --diff-filter=ACMR 2>/dev/null \
      | grep -E '^src/(infra/.*\.(ts|tsx)|main/java/.*infra.*\.java)$' || true
  fi
}

# Prints the original line number of the first fetch call with no deadline
# marker within the 8 lines that follow it (comment lines dropped first), or
# nothing when every call is covered.
first_uncovered_fetch() {
  awk '
    /^[[:space:]]*(\/\/|\*|\/\*)/ { next }
    { orig[++n] = NR; line[n] = $0 }
    END {
      for (i = 1; i <= n; i++) {
        if (line[i] ~ /(^|[^A-Za-z_.$])fetch\(/ || line[i] ~ /globalThis\.fetch\(/) {
          ok = 0
          for (j = i; j <= i + 8 && j <= n; j++) if (line[j] ~ /AbortSignal\.timeout\(|signal:/) ok = 1
          if (!ok) { print orig[i]; exit }
        }
      }
    }'
}

status=0
while IFS= read -r f; do
  [ -n "$f" ] || continue
  case "$f" in *.test.ts|*.test.tsx|*test-helpers*) continue ;; esac
  # Staged mode reads what will be committed, the index, never the working tree: a
  # deadline added in the editor but left unstaged passed the gate until 2026-09-27.
  if [ "$MODE" = "--all" ]; then
    [ -f "$f" ] || continue
    content=$(cat "$f")
  else
    content=$(git show ":$f")
  fi
  case "$f" in
    *.java)
      if echo "$content" | grep -q 'HttpClient\.new' \
        && ! echo "$content" | grep -qE '\.timeout\(|connectTimeout'; then
        echo "  ╳ $f opens an HttpClient with no timeout (rule 29)" >&2
        status=1
      fi
      ;;
    *)
      uncovered=$(printf '%s\n' "$content" | first_uncovered_fetch)
      if [ -n "$uncovered" ]; then
        echo "  ╳ $f:$uncovered calls fetch with no deadline marker within 8 lines (rule 29)" >&2
        status=1
      fi
      ;;
  esac
done < <(candidate_files)

[ "$status" -eq 0 ] || echo "  fix: AbortSignal.timeout(ms) / client timeouts on every outbound call (references/reliability.md)" >&2
exit "$status"
