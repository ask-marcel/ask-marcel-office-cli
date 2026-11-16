# Changelog

All notable changes to `ask-marcel-office-cli` from 3.0.0 are documented here.
The history of 2.8.0 and before is in the root `CHANGELOG.md`.

## Unreleased

### 3.0.0: the package is a facade of the three packages

`ask-marcel-office-cli` 3.0.0 holds no command. It installs
`@ask-marcel/office-auth`, `@ask-marcel/office-read` and
`@ask-marcel/office-write`, each at one exact version, and their three bins
([ADR 0005](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/docs/adr/0005-facade-package-ask-marcel-office-cli-3.md)).

If you used `ask-marcel-office-cli` 2.8.0, these changes apply to you:

- The bin `ask-marcel-office` is gone (a breaking change). The bins are
  `ask-marcel-office-auth`, `ask-marcel-office-read` and
  `ask-marcel-office-write`.
- `login`, `logout` and `status` are commands of `ask-marcel-office-auth`.
- The mail drafts and the PDF conversions are commands of
  `ask-marcel-office-write`. All the other commands are commands of
  `ask-marcel-office-read`.
- `scopes-check` and `update` are gone. Use `ask-marcel-office-auth status`,
  and update with `npm i -g ask-marcel-office-cli@latest`.
- The package has no library entry. Import the library of a package by its
  name, for example `@ask-marcel/office-read`.
- There is no combined MCP server. Register the MCP server of each package
  that you use.

The changelog of each package gives the changes of its commands since 2.8.0.
