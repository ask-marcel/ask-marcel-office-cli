# Contributing

This guide is for a person who changes the code of this repository. To use the
packages, read the README of each package:
[auth](auth/README.md), [read](read/README.md) and [write](write/README.md).

## The workspace

The repository is one Bun workspace with four folders. The root holds no
package and no `src/`. It keeps the shared configs, `scripts/` and the docs.

| Folder | npm name | Published | What it holds |
|:--|:--|:--|:--|
| `core/` | `@ask-marcel/office-core` | No (private) | The shared code: the domain types, the Graph clients, the CLI and MCP shells, the presenter. Each bundle inlines it. |
| `auth/` | `@ask-marcel/office-auth` | Yes | The sign-in, the token cache and the token helper |
| `read/` | `@ask-marcel/office-read` | Yes | The read commands and the converters |
| `write/` | `@ask-marcel/office-write` | Yes | The mail drafts and the PDF conversions |

Each folder has its own `package.json`, configs, tests and gates, and the same
layout: `src/{domain,use-cases,infra,presenter,composition}`. The folder of a
file decides its package (`scripts/check-package-imports.ts`). The plan of the
split is [docs/plans/2026-10-01-package-split.md](docs/plans/2026-10-01-package-split.md),
and the decision is [ADR 0003](docs/adr/0003-split-into-auth-read-write-packages.md).

Set up a checkout one time:

```sh
bun install
git config core.hooksPath .githooks
```

In a new checkout, run the auth bin one time (`bun run src/main.ts status` in
`auth/`) before a read or write command. Only the auth bin records the token
helper in `~/.ask-marcel/token-helper.json`. Run it again when you change
between the source and `dist/`, because the file names the last one that ran.

## The commands of each folder

Run these commands in the folder that you change:

| Command | What it does | Folders |
|:--|:--|:--|
| `bun test` | The tests of the folder | all |
| `bun run lint` / `bun run lint:strict` | ESLint (fast / type-aware), 0 errors and 0 warnings | all |
| `bun run typecheck` | `tsc --noEmit` | all |
| `bun run coverage` | 100% on every tier of `src/` | all |
| `bun run mutate:changed` | Stryker on the changed domain and use-case files (90% or more) | all |
| `bun run build` | `dist/`: the bundles with core inlined, and the declarations | auth, read, write |
| `bun run check:packed` | Packs the package, and installs the tarball in a clean consumer | auth, read, write |
| `bun run docs:gen` | Makes the generated tables of `COMMANDS.md` (and, in `build`, `dist/commands.json`) | auth, read, write |
| `bun run check:docs` | Fails when `COMMANDS.md` is out of date. In read and write, it also fails when `dist/commands.json` is out of date. | auth, read, write |
| `bun run smoke` | Runs `dist/` under node and bun | auth, read, write |
| `bun run smoke:packed` | Installs the packed tarball in an empty folder and runs the smoke | read |

At the root, `bun run test` runs the tests of every folder. `bun run lint` lints
the root config files, and `bun run typecheck` checks `scripts/`.
`bun run check:read-bundle` and `bun run check:package-imports` check the
package borders.

When you change the registry of a package, run `bun run docs:gen` in its
folder and commit the new `COMMANDS.md`. Do not edit the generated tables by
hand.

## How the code is built

`core/src/composition/cli-shell.ts` (commander) and
`core/src/composition/mcp-shell.ts` (MCP over stdio) both call
`core/src/composition/run-registry-command.ts`. That module does all the work
between "a command and its params" and "a value or a failure". It maps the
option names, reads the local files and sets the source of an error. It also
does `--output-path` and `--output-dir`. A fix there reaches both front ends.

These shells and the docs builders import no registry. They get the package
(its registry, its lifecycle entries, its bin and its package name) as an
argument. Each package gives its own: `read/src/composition/single-package.ts`,
`write/src/composition/write-package.ts` and
`auth/src/composition/auth-package.ts`.

`core/src/presenter/render-to-string.ts` makes the output string, and
`core/src/presenter/output.ts` is the only module that writes to stdout. The MCP
server cannot write to stdout, because stdout carries its JSON-RPC frames. Thus
it uses the string and gives it back as the tool content.

## The pre-commit hook

`.githooks/pre-commit` has five fast gates: the commit size (10 files and 300
lines at most), `package.json` (no `latest` or `*`), gitleaks, the staged lint
and the typecheck. The lint and the typecheck run in each folder with a staged
change. A change under `core/` also runs them in `auth/`, `read/` and `write/`.
A change outside the folders runs them at the root.

Install gitleaks, or the hook skips the secret scan and says so only in its
log. `.githooks/commit-msg` refuses a header that is not a Conventional Commit:
`type(scope)!: subject`, with a type from
`feat fix docs style refactor perf test build ci chore revert`, a lowercase
scope (no comma) and no period at the end, 100 characters at most.

The hook does not run the tests, the coverage, `lint:strict` or the mutation.
Run them in the folder before you commit. CI runs them all.

## CI

`.github/workflows/ci.yml` is the gate that you cannot skip. Its jobs:

| Job | What it runs |
|:--|:--|
| `gates` | The scripts and their selftests, the commit messages, the secret scan of the full history, the root lint and typecheck, the package borders, the red fixture of the command reference and the dependency audit |
| `workspace` | One job for each folder: `lint:strict`, typecheck, tests, coverage, the mutation of the changed files (not in `read/`), and for a published folder `build`, `check:docs` and `check:packed`. Auth and write also run their bundle smoke here. |
| `mutation` | The mutation of the changed files of `read/`, in 8 shards |
| `smoke` | The build of `read/`, its bundle smoke and its packed-install smoke |
| `windows-helper` | The token helper round trip on Windows |

`gates` fails when the `workspace` matrix misses a folder
(`scripts/check-ci-matrix.ts`).

## The docs

- Each published folder has `README.md` (for npm), `USAGE.md` and
  `COMMANDS.md`. A doc gives only the counts of its own package.
- The root `README.md` is the index of the packages.
- `docs/COMMANDS.md`, `docs/USAGE.md`, `docs/commands.json` and `docs/demo.gif`
  are the docs of `ask-marcel-office-cli` 2.8.0. Its npm page links them. Do
  not change them for the packages.
- [docs/QA-PLAYBOOK.md](docs/QA-PLAYBOOK.md) is the full health check of the
  command surface. Run it before a release.
- The decisions are in [docs/adr/](docs/adr/), and the lessons are in
  `.claude/LESSONS.md`.
