#!/usr/bin/env bash
#
# The mutation scope, shared by mutate-changed.sh and mutate-staged.sh and
# proven by mutation-scope-selftest.sh. Source it; it defines three functions:
#
#   scope_changed_paths <git diff args>  repo-root paths whose diff is more than
#                                        import lines (renames print the new path)
#   scope_untracked                      repo-root paths of untracked files
#   scope_filter                         stdin repo-root paths -> stdout the files
#                                        to mutate, relative to the working dir
#
# Two properties the package split depends on:
#
#   1. A file whose every changed line is an import or re-export statement is
#      out of scope. Moving a file and fixing its imports changes no logic, and
#      without this a mass move hands CI the whole mutation scope, which cannot
#      finish within the job limit. `git diff -I` ignores hunks whose every line
#      matches; git then omits the file or reports it 0 added / 0 deleted, and
#      either way it is dropped. The cost is accepted knowingly: an import-only
#      edit can still change behaviour (a swapped binding, a dropped side-effect
#      import), and such an edit is not mutated.
#   2. git prints paths from the repo root, so a run from a sub-folder (one
#      package of a workspace) maps them through `git rev-parse --show-prefix`.
#      Filtering unmapped paths on `^src/` matched nothing there and passed green.
#
# Stage a move (git mv, or git add -A) before running mutate:changed: to git an
# unstaged move is a new untracked file, so it is mutated in full.
#
# Every function returns git's status (callers run under pipefail), so a diff
# that fails fails the run instead of reading as "nothing changed".
#
# Bash 3.2 compatible (the macOS system bash).

# One import, export-from or side-effect import statement alone on its line,
# with optional import attributes, semicolon and trailing comment. No quote may
# come before ` from `, so an exported string that merely contains ` from '...'`
# is logic and stays in scope. POSIX ERE without \b: git's -I uses the platform
# regex, and macOS has no \b.
SCOPE_IMPORT_ONLY_RE="^((import|export) [^'\"]* from ['\"][^'\"]*['\"]|} from ['\"][^'\"]*['\"]|import ['\"][^'\"]*['\"])([[:space:]]+with[[:space:]]*[{][^}]*[}])?;?[[:space:]]*(//.*)?$"

scope_toplevel() {
  git rev-parse --show-toplevel
}

# Reads `git diff --numstat -z` records; prints the path of every non-0/0 row.
# A rename record is "added<TAB>deleted<TAB>" followed by two NUL-terminated
# paths, old then new.
scope_parse_numstat() {
  local rec rest added deleted path old
  while IFS= read -r -d '' rec; do
    added=${rec%%$'\t'*}
    rest=${rec#*$'\t'}
    deleted=${rest%%$'\t'*}
    path=${rest#*$'\t'}
    if [ -z "$path" ]; then
      IFS= read -r -d '' old
      IFS= read -r -d '' path
    fi
    if [ "$added" = 0 ] && [ "$deleted" = 0 ]; then
      continue
    fi
    printf '%s\n' "$path"
  done
}

scope_changed_paths() {
  git -C "$(scope_toplevel)" diff -M --numstat -z --diff-filter=ACMR -I"$SCOPE_IMPORT_ONLY_RE" "$@" | scope_parse_numstat
}

scope_untracked() {
  git -C "$(scope_toplevel)" ls-files --others --exclude-standard
}

# A changed test file pulls in the source it covers: test files carry no
# mutants, so a commit touching only tests would otherwise mutate nothing and
# hide a weakened test. Partial by construction: a shared test file with no
# sibling source maps to nothing, and the existence check drops those misses
# (and files deleted in the working tree).
scope_filter() {
  local prefix p f
  prefix=$(git rev-parse --show-prefix)
  while IFS= read -r p; do
    case "$p" in
      "$prefix"*) printf '%s\n' "${p#"$prefix"}" ;;
    esac
  done \
    | sed -n -E -e 's/\.test\.ts$/.ts/' -e '/\/ports\//d' -e '/^src\/(domain|use-cases)\/.*\.ts$/p' \
    | sort -u \
    | while IFS= read -r f; do
        if [ -f "$f" ]; then printf '%s\n' "$f"; fi
      done
}
