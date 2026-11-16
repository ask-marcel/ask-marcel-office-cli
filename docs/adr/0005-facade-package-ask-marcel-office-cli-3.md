# 0005: A facade package, `ask-marcel-office-cli` 3.x

- Status: accepted (2026-11-16)
- Date: 2026-11-16
- Changes: ADR 0003's consequence "`ask-marcel-office-cli` is frozen at 2.8.0 and deprecated once
  the three 1.0.0 releases are live", and decision D12 of
  `docs/plans/2026-10-01-package-split.md`, which refused a facade meta-package

## Context

ADR 0003 split `ask-marcel-office-cli` 2.8.0 into `@ask-marcel/office-auth`,
`@ask-marcel/office-read` and `@ask-marcel/office-write`. A host that wants only the read commands
installs only read, and that need stays.

A person who wants all the commands on one computer has a different need: one install, one version
to name, and the old package gone. Three installs with three versions do not give that. The old
name also has users, and a deprecation note alone gives them no upgrade under the name that they
know.

## Decision

A fifth workspace folder, `cli/`, holds the package `ask-marcel-office-cli` at 3.0.0:

- It depends on the three packages, each at one exact version. In the repo the dependencies are
  `workspace:*`, and `bun pm pack` writes the exact versions.
- It holds no command, no registry and no library entry. The three packages keep all the code.
- It installs three bins with the names of the three packages: `ask-marcel-office-auth`,
  `ask-marcel-office-read` and `ask-marcel-office-write`. A package manager links only the bins of
  the package that a person installs, so the facade has its own bin files. Each one finds the
  `package.json` of its package, reads the `bin` entry and starts that file in the same process.
- It has no `ask-marcel-office` bin. This is the breaking change of 3.0.0.
- `cli/` is a full workspace folder: the same layout, tests, coverage and CI job as the other
  folders.
- The root `package.json` gets a private name of its own, because two workspace packages cannot
  have the same name.

The three packages stay as ADR 0003 decided: they release alone, and a host can install read
without write. The facade adds a fourth way to install, it does not replace the three.

Until a release to npm, the facade installs from local tarballs (a script in `scripts/`). A release
to npm is a different decision, with the QA playbook before it.

## Options considered

- **A local facade outside the repo.** Refused by the owner: the combined package must be a part of
  the product, with its gates.
- **A new scoped name, `@ask-marcel/office`.** Refused by the owner: the old name gives the users of
  2.8.0 an upgrade under the name that they know.
- **One merged package with one bin.** Refused: it reverses ADR 0003, and an install then always
  holds the write commands.
- **A fourth bin `ask-marcel-office` that sends a command to its package.** Refused: `mcp`, `docs`
  and `help-json` exist in each of the three packages, so the bin cannot route them without one
  merged registry, and each error hint already names the bin of its package.
- **Depend on a subpath export of each bin.** Refused: `dist/cli.js` then becomes a public entry of
  the three packages, and an import of it runs the bin.

## Consequences

- A script or an agent that calls `ask-marcel-office` stops after the upgrade to 3.0.0. The
  changelog of the facade names the three bins.
- The facade does not make a combined MCP server. A host registers up to three servers, as in ADR
  0003.
- An install of the facade always holds the write commands. A host that must never write installs
  read and auth, not the facade.
- A release of the facade names one tested set of the three versions. A fix in one package reaches
  the users of the facade only after a new facade version.
- The update notice of a package bin names its own package, not the facade. A person who installed
  the facade then reads an advice to install one package. A later change can make the notice name
  the facade.
- `docs/COMMANDS.md`, `docs/USAGE.md`, `docs/commands.json` and `docs/demo.gif` stay the frozen
  docs of 2.8.0, because the npm page of 2.8.0 links them.
- The gate scripts that list the workspace folders get the new folder.

## Reversal

Before a release to npm: remove `cli/`, its lines in the gate scripts and the CI matrix, and give
the root package its old name. After a release: deprecate the 3.x versions with a note that names
the three packages. The three packages do not depend on the facade, so nothing else changes.
