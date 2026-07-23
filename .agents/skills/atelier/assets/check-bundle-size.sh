#!/usr/bin/env bash
#
# Bundle weight budget (rule 17.7): a light interface is a light payload, so make
# it a number the pipeline enforces, not an adjective nobody measures. Fail the
# build when the built JS crosses its gzipped ceiling.
#
# Point it at the built output directory (a Next static export's `out/`, a
# bundler's `dist/`) and a budget in gzipped KB. Measure on the shipped bundle,
# and treat a bump as a deliberate, reviewed change, not a silent drift.
#
#   BUDGET_KB=180 bash scripts/check-bundle-size.sh out
#
set -euo pipefail

budget_kb="${BUDGET_KB:-180}"
dir="${1:-out}"

if [ ! -d "$dir" ]; then
  echo "check-bundle-size: '$dir' not found (build the app first?)" >&2
  exit 1
fi

total=0
found=0
while IFS= read -r f; do
  found=1
  sz=$(gzip -c "$f" | wc -c | tr -d ' ')
  total=$((total + sz))
done < <(find "$dir" -type f -name '*.js')

if [ "$found" -eq 0 ]; then
  echo "check-bundle-size: no .js files under '$dir' (build the app first?)" >&2
  exit 1
fi

total_kb=$(( (total + 1023) / 1024 ))
echo "check-bundle-size: ${total_kb} KB gzipped in ${dir} (budget ${budget_kb} KB)"

if [ "$total_kb" -gt "$budget_kb" ]; then
  echo "  OVER BUDGET by $((total_kb - budget_kb)) KB (rule 17.7). Trim the bundle, or raise the budget deliberately in the CI config with a reason." >&2
  exit 1
fi

echo "  within budget"
