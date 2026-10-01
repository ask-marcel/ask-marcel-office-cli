#!/usr/bin/env bash
#
# Proves the mutation scope (scripts/mutation-scope.sh) and its two callers map
# a diff to the right files, so `mutate:changed` and `mutate:staged` can neither
# explode on a mass move nor pass having mutated nothing. Each case builds a
# throwaway repo, so the real history is never read; stryker is replaced by a
# fake `bunx` that records the files it was asked to mutate.
#
# Why the cases exist:
#   - a move with only import edits mutates nothing: the package split moves
#     hundreds of files, and CI cannot mutate them all within its job limit;
#   - a logic edit inside a moved file is still mutated, and so is an exported
#     string that merely reads like an import: the import filter must not hide
#     real changes;
#   - run from a sub-folder, the repo-root paths git prints map to that folder:
#     a `^src/` filter on unmapped paths matched nothing and passed green;
#   - a diff that fails (no merge base) fails the run instead of reading as
#     "nothing changed".
#
# Run: bash scripts/mutation-scope-selftest.sh

set -euo pipefail

here=$(cd "$(dirname "$0")" && pwd)
# shellcheck source=scripts/mutation-scope.sh
. "$here/mutation-scope.sh"

failures=0
work=$(mktemp -d "${TMPDIR:-/tmp}/mutation-scope-selftest.XXXXXX")
trap 'rm -rf "$work"' EXIT

mkdir -p "$work/bin"
printf '#!/usr/bin/env bash\nwhile [ $# -gt 0 ]; do if [ "$1" = --mutate ]; then echo "$2" > "$STRYKER_ARGS"; fi; shift; done\n' > "$work/bin/bunx"
chmod +x "$work/bin/bunx"

expect() {
  if [ "$2" = "$3" ]; then
    echo "  ok    $1"
  else
    echo "  FAIL  $1"
    echo "        want: [$(echo "$2" | paste -sd, -)]"
    echo "        got:  [$(echo "$3" | paste -sd, -)]"
    failures=$((failures + 1))
  fi
}

new_repo() {
  mkdir -p "$work/$1/src/use-cases/ports" "$work/$1/src/domain" "$work/$1/pkg/src/use-cases"
  cd "$work/$1"
  git init -q
  git config user.email selftest@example.invalid
  git config user.name selftest
  # On by default, but a config can turn it off: the helper's explicit -M must carry it.
  git config diff.renames false
  printf "import { a } from './a.ts';\nimport {\n  b,\n} from './b.ts';\nexport const f = (): number => a + b;\nexport const hint = 'run it from \"the repo root\"';\n" > src/use-cases/x.ts
  for n in 1 2 3 4 5 6; do printf "export const f%s = (): number => %s;\n" "$n" "$n" >> src/use-cases/x.ts; done
  printf "export const g = (): number => 2;\n" > src/use-cases/y.ts
  printf "test('g', () => {});\n" > src/use-cases/y.test.ts
  printf "export type P = () => number;\n" > src/use-cases/ports/p.ts
  printf "export const d = (): number => 4;\n" > src/domain/d.ts
  printf "export const z = (): number => 5;\n" > pkg/src/use-cases/z.ts
  git add -A
  git commit -qm base
}

# Moves x.ts to $1 with its import specifiers rewritten, plus the sed edit $2.
move_x() {
  mkdir -p "$(dirname "$1")"
  git mv src/use-cases/x.ts "$1"
  sed -e "s#'./a.ts'#'../a.ts'#" -e "s#} from './b.ts'#} from '../b.ts'#" -e "${2:-}" "$1" > x.tmp
  mv x.tmp "$1"
  git add -A
}

edit_commit() {
  sed -e "$2" "$1" > x.tmp && mv x.tmp "$1"
  git add -A && git commit -qm edit
}

committed_scope() {
  scope_changed_paths HEAD~1...HEAD | scope_filter
}

rename_status() {
  git diff -M --name-status HEAD~1 HEAD | cut -c1 | sort -u | paste -sd, -
}

echo "mutation-scope selftest:"

new_repo move-imports-only
move_x src/use-cases/sub/x.ts && git commit -qm move
expect 'precondition: the import-only move is a rename' 'R' "$(rename_status)"
expect 'a move with only import edits mutates nothing' '' "$(committed_scope)"

new_repo move-with-logic
move_x src/use-cases/sub/x.ts 's/a + b/a - b/' && git commit -qm move
expect 'precondition: the logic move is a rename' 'R' "$(rename_status)"
expect 'a logic edit inside a moved file is mutated' 'src/use-cases/sub/x.ts' "$(committed_scope)"

new_repo staged-move
move_x src/use-cases/sub/x.ts
expect 'a staged pure move mutates nothing' '' "$(scope_changed_paths --cached | scope_filter)"

new_repo export-string
edit_commit src/use-cases/x.ts 's/the repo root/the package root/'
expect 'an exported string reading like an import is logic' 'src/use-cases/x.ts' "$(committed_scope)"

new_repo tests-and-ports
printf "test('g', () => { expect(1).toBe(1); });\n" > src/use-cases/y.test.ts
printf "export type P = () => string;\n" > src/use-cases/ports/p.ts
printf "test('shared', () => {});\n" > src/use-cases/shared.test.ts
edit_commit src/domain/d.ts 's/4/6/'
expect 'a changed test pulls in its source (none for a shared test); ports are never mutated' "$(printf 'src/domain/d.ts\nsrc/use-cases/y.ts')" "$(committed_scope)"

new_repo sub-folder
git config diff.relative true
edit_commit pkg/src/use-cases/z.ts 's/5/7/'
expect 'a run at the root ignores another folder' '' "$(committed_scope)"
expect 'a run from the folder maps repo-root paths to it' 'src/use-cases/z.ts' "$(cd pkg && committed_scope)"

new_repo cross-folder-move
move_x pkg/src/use-cases/x.ts && git commit -qm move
expect 'a pure move into another folder mutates nothing there' '' "$(cd pkg && committed_scope)"
edit_commit pkg/src/use-cases/x.ts 's/a + b/a * b/'
expect 'a logic edit after the move is mutated from its folder only' 'src/use-cases/x.ts' "$(cd pkg && committed_scope)"
expect 'and not from the root' '' "$(committed_scope)"

new_repo untracked
printf "export const n = (): number => 8;\n" > src/domain/n.ts
printf "export const m = (): number => 9;\n" > pkg/src/use-cases/m.ts
expect 'an untracked source is mutated from the root' 'src/domain/n.ts' "$(scope_untracked | scope_filter)"
expect 'an untracked source is mutated from its folder' 'src/use-cases/m.ts' "$(cd pkg && scope_untracked | scope_filter)"

new_repo callers
edit_commit src/use-cases/y.ts 's/2/3/'
export STRYKER_ARGS="$work/stryker-args"
rm -f "$STRYKER_ARGS"
PATH="$work/bin:$PATH" MUTATE_NO_FETCH=1 BASE=HEAD~1 bash "$here/mutate-changed.sh" > /dev/null
expect 'mutate:changed hands stryker the changed file' 'src/use-cases/y.ts' "$(cat "$STRYKER_ARGS" 2>/dev/null || true)"
printf "export const u = (): number => 1;\n" > src/domain/u.ts
PATH="$work/bin:$PATH" MUTATE_NO_FETCH=1 BASE=HEAD bash "$here/mutate-changed.sh" > /dev/null
expect 'mutate:changed hands stryker an untracked file' 'src/domain/u.ts' "$(cat "$STRYKER_ARGS" 2>/dev/null || true)"
rm -f src/domain/u.ts
sed -e 's/4/7/' src/domain/d.ts > x.tmp && mv x.tmp src/domain/d.ts
PATH="$work/bin:$PATH" MUTATE_NO_FETCH=1 BASE=HEAD bash "$here/mutate-changed.sh" > /dev/null
expect 'mutate:changed hands stryker an unstaged edit' 'src/domain/d.ts' "$(cat "$STRYKER_ARGS" 2>/dev/null || true)"
git checkout -q -- src/domain/d.ts
unrelated=$(git commit-tree "$(git hash-object -t tree /dev/null)" -m unrelated)
status=0
PATH="$work/bin:$PATH" MUTATE_NO_FETCH=1 BASE="$unrelated" bash "$here/mutate-changed.sh" > /dev/null 2>&1 || status=$?
expect 'mutate:changed fails when a diff fails (no merge base)' 'failed' "$([ "$status" -ne 0 ] && echo failed || echo passed)"
rm -f "$STRYKER_ARGS"
edit_commit src/domain/d.ts 's/4/5/'
git reset -q --soft HEAD~1
PATH="$work/bin:$PATH" bash "$here/mutate-staged.sh" > /dev/null
expect 'mutate:staged hands stryker the staged file' 'src/domain/d.ts' "$(cat "$STRYKER_ARGS" 2>/dev/null || true)"

if [ "$failures" -gt 0 ]; then
  echo "mutation-scope selftest: ${failures} case(s) failed" >&2
  exit 1
fi
echo "mutation-scope selftest: passed"
