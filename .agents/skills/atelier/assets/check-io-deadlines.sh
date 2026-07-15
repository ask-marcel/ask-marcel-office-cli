#!/usr/bin/env bash
#
# Rule 29 tripwire: every outbound call in an infra adapter has a deadline.
#
# Checks files the STAGED DIFF touches under the infra layer; `--all` scans the
# whole layer (adopt-mode audit). Per-file heuristic: a file that makes an
# outbound call must also mention a deadline marker somewhere in the file.
#
#   TS   call marker:      fetch(
#        deadline markers: AbortSignal | signal: | timeout
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

status=0
while IFS= read -r f; do
  [ -n "$f" ] && [ -f "$f" ] || continue
  case "$f" in *.test.ts|*.test.tsx|*test-helpers*) continue ;; esac
  content=$(cat "$f")
  case "$f" in
    *.java)
      if echo "$content" | grep -q 'HttpClient\.new' \
        && ! echo "$content" | grep -qE '\.timeout\(|connectTimeout'; then
        echo "  ╳ $f opens an HttpClient with no timeout (rule 29)" >&2
        status=1
      fi
      ;;
    *)
      if echo "$content" | grep -qE '(^|[^.a-zA-Z])fetch\(' \
        && ! echo "$content" | grep -qE 'AbortSignal|signal:|timeout'; then
        echo "  ╳ $f calls fetch with no deadline marker (rule 29)" >&2
        status=1
      fi
      ;;
  esac
done < <(candidate_files)

[ "$status" -eq 0 ] || echo "  fix: AbortSignal.timeout(ms) / client timeouts on every outbound call (references/reliability.md)" >&2
exit "$status"
