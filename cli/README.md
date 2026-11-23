# ask-marcel-office-cli

All the ask-marcel-office commands in one install. This package holds no
command of its own: it installs the three packages of the project and their
three bins.

| Bin | Package | What it does |
|:--|:--|:--|
| `ask-marcel-office-auth` | [`@ask-marcel/office-auth`](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/README.md) | Signs you in, keeps the token cache and gives the tokens to the other two bins |
| `ask-marcel-office-read` | [`@ask-marcel/office-read`](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/README.md) | Reads your Microsoft 365 data, and converts files and mail to markdown |
| `ask-marcel-office-write` | [`@ask-marcel/office-write`](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/README.md) | Writes mail drafts and calendar drafts that it never sends, changes the state of mail and the tasks of To Do and Planner, and converts attachments to PDF |

Each bin is the bin of its package, with the same commands, the same output and
the same exit codes. The docs of each package apply with no change.

## Install

```sh
npm i -g ask-marcel-office-cli
ask-marcel-office-auth login
ask-marcel-office-read my-quick-context
```

The package needs Node 20 or later, or Bun 1.0 or later, on macOS, Windows or
Linux.

## When to install the packages one by one

An install of this package always holds the write commands. For a host that
must never write, install `@ask-marcel/office-auth` and
`@ask-marcel/office-read`, not this package.

Each package registers its own MCP server (`ask-marcel-office-read mcp`, for
example). This package adds no combined server.

## If you used 2.8.0

The bin `ask-marcel-office` is gone. Use the bin of the package that owns the
command: `ask-marcel-office-auth` for `login`, `logout` and `status`,
`ask-marcel-office-write` for the mail drafts and the PDF conversions, and
`ask-marcel-office-read` for all the other commands. The write bin also has
commands that 2.8.0 did not have. The
[changelog](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/cli/CHANGELOG.md)
gives the full list of changes.
