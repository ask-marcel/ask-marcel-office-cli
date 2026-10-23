#!/usr/bin/env bash
#
# Prints, one per line, the folders whose fast checks the pre-commit hook runs
# for the staged change (package split, D7: the root hook is the one
# dispatcher). A folder is the root (".") or one of the workspace folders that
# the root package.json lists (core, auth, read, write).
#
#   - a workspace folder runs when a staged path is under it;
#   - a staged path under core/ runs every workspace folder, since auth, read
#     and write all build on core; so does a change to a root config that every
#     folder extends (tsconfig.json, eslint.base.js) or to the ownership lists
#     that the folder configs import (eslint.ownership.js);
#   - the root runs when a staged path is outside every workspace folder;
#   - a move counts on both sides: `git mv src/x.ts core/src/x.ts` runs the
#     root as well as every folder.
#
#   --selftest  proves each rule on canned staged paths (rule 15.10)
#
# Bash 3.2 compatible (the macOS system bash).

set -euo pipefail

SHARED_CONFIGS_RE='^(tsconfig\.json|eslint\.base\.js|eslint\.ownership\.js)$'

workspace_folders() {
  bun --print "((await Bun.file('package.json').json()).workspaces ?? []).join('\n')"
}

# The repo-root paths of the staged change. --no-renames lists both sides of
# a move (the old path as deleted, the new one as added), so the folder that a
# file leaves is checked as well as the folder that it enters.
staged_paths() {
  git diff --cached --name-only --no-renames --diff-filter=ACMRD
}

# stdin: staged repo-root paths. $1: the workspace folders, one per line.
dispatch() {
  local folders="$1" path folder all="" root="" picked=""
  while IFS= read -r path; do
    [ -n "$path" ] || continue
    case "$path" in core/*) all=1 ;; esac
    if printf '%s\n' "$path" | grep -qE "$SHARED_CONFIGS_RE"; then all=1; fi
    folder=${path%%/*}
    if [ "$folder" != "$path" ] && printf '%s\n' "$folders" | grep -qxF "$folder"; then
      picked="${picked}${folder}"$'\n'
    else
      root=1
    fi
  done
  if [ -n "$root" ]; then echo "."; fi
  if [ -n "$all" ]; then
    printf '%s\n' "$folders"
    return 0
  fi
  printf '%s' "$picked" | awk 'NF && !seen[$0]++'
}

selftest() {
  local folders status=0
  folders=$'core\nauth\nread\nwrite'
  # expect <label> <expected, space-separated> <staged paths, space-separated>
  expect() {
    local got
    got=$(printf '%s\n' $3 | dispatch "$folders" | tr '\n' ' ' | sed 's/ $//')
    if [ "$got" != "$2" ]; then
      echo "staged-folders: SELFTEST FAILED, $1: expected '$2', got '$got'" >&2
      status=1
    fi
  }
  expect 'a root source alone runs the root' '.' 'src/main.ts'
  expect 'a read source alone runs read' 'read' 'read/src/index.ts'
  expect 'two folders run each once' 'auth write' 'auth/src/a.ts write/src/b.ts auth/src/c.ts'
  expect 'a core change runs every workspace folder' 'core auth read write' 'core/src/domain/result.ts'
  expect 'a root config every folder extends runs them all, and the root' '. core auth read write' 'tsconfig.json'
  expect 'the ownership lists that the folder configs import run them all, and the root' '. core auth read write' 'eslint.ownership.js'
  expect 'a root file and a folder file run both' '. write' 'docs/plans/x.md write/README.md'
  expect 'a root file that only looks like a folder is the root' '.' 'readme-notes.md'
  # expect_move <label> <expected> <from> <to>: a real `git mv` in a scratch
  # repo, so the staged listing itself is under test, not only the dispatch.
  expect_move() {
    local repo got
    repo=$(mktemp -d)
    (
      cd "$repo"
      git init -q
      mkdir -p "$(dirname "$3")" "$(dirname "$4")"
      printf 'export const moved = 1;\n' >"$3"
      git add -A
      git -c user.name=selftest -c user.email=selftest@example.invalid -c core.hooksPath=/dev/null -c commit.gpgsign=false commit -qm seed
      git mv "$3" "$4"
    )
    got=$(cd "$repo" && staged_paths | dispatch "$folders" | tr '\n' ' ' | sed 's/ $//')
    rm -rf "$repo"
    if [ "$got" != "$2" ]; then
      echo "staged-folders: SELFTEST FAILED, $1: expected '$2', got '$got'" >&2
      status=1
    fi
  }
  expect_move 'a move out of the root src into core runs the root and every folder' '. core auth read write' 'src/x.ts' 'core/src/x.ts'
  expect_move 'a move from read to write runs both' 'read write' 'read/src/a.ts' 'write/src/a.ts'
  if [ "$status" -eq 0 ]; then
    echo "staged-folders: selftest passed (a folder, two folders, core, a shared config, the root and both sides of a move each dispatch as promised)."
  fi
  return "$status"
}

if [ "${1:-}" = "--selftest" ]; then
  selftest
  exit $?
fi

staged_paths | dispatch "$(workspace_folders)"
