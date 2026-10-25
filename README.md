<div align="center">

# ask-marcel-office

### The Microsoft 365 command line built for AI agents

**One sign-in. No Azure app registration. Every document as clean markdown.**

Outlook · OneDrive · SharePoint · Calendar · Excel · Teams · Planner · To Do · OneNote · People

[Packages](#the-three-packages) · [Install](#install) · [Agent skill](#the-agent-skill) · [Develop](#develop)

</div>

---

ask-marcel-office gives a tool-calling agent (Claude Code, Cursor, Cline, an MCP
host or your own loop) access to your Microsoft 365 data through Microsoft
Graph. You sign in one time, in a browser, as yourself. You do not need an Azure
app registration, the consent of a tenant admin or a client secret. Office
files, PDF files and mail become clean markdown, so the agent can read them.

The project is three npm packages. Each package has its own bin, its own MCP
server, its own library and its own docs.

## The three packages

| Package | Bin | What it does | Do you need it? | Docs |
|:--|:--|:--|:--|:--|
| `@ask-marcel/office-auth` | `ask-marcel-office-auth` | Signs you in, keeps the token cache and gives the tokens to the other two packages. | Yes | [README](auth/README.md) · [USAGE](auth/USAGE.md) · [COMMANDS](auth/COMMANDS.md) |
| `@ask-marcel/office-read` | `ask-marcel-office-read` | Reads mail, calendar, OneDrive and SharePoint, Teams and chats, users, Excel, OneNote, To Do and Planner. Converts files and mail to markdown. It cannot change your tenant. | Yes | [README](read/README.md) · [USAGE](read/USAGE.md) · [COMMANDS](read/COMMANDS.md) |
| `@ask-marcel/office-write` | `ask-marcel-office-write` | Writes mail drafts and never sends them. Converts mail, group post and calendar attachments to PDF through a temporary OneDrive file. | Only for drafts and PDF conversions | [README](write/README.md) · [USAGE](write/USAGE.md) · [COMMANDS](write/COMMANDS.md) |

Start with the [read package README](read/README.md). It shows what an agent
can do with your Microsoft 365 data.

## Install

> The three packages are not on npm yet. npm installs `ask-marcel-office-cli`
> 2.8.0 now. Its docs are [docs/USAGE.md](docs/USAGE.md) and
> [docs/COMMANDS.md](docs/COMMANDS.md).

To read only, install auth and read, then sign in:

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
ask-marcel-office-read my-quick-context
```

To also write mail drafts and convert attachments to PDF, install the three
packages:

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read @ask-marcel/office-write
```

The packages need Node 20 or later, or Bun 1.0 or later, on macOS, Windows or
Linux. Each package README shows how to register its MCP server.

## The agent skill

The [`skills/`](skills/SKILL.md) folder holds an agent skill: it tells an agent
how to answer a question from your Microsoft 365 data, and how to write a reply
as an unsent draft. No package ships it. To use it, copy `skills/` to
`~/.claude/skills/ask-marcel/`, or to the skill folder of your agent.

## Develop

- [CONTRIBUTING.md](CONTRIBUTING.md): the workspace, the commands of each
  folder, the pre-commit hook and the CI jobs.
- [QA playbook](docs/QA-PLAYBOOK.md): the full health check of the command
  surface before a release.
- [The package split plan](docs/plans/2026-10-01-package-split.md) and the
  [architecture decision records](docs/adr/).
- Questions, requests and pull requests: the
  [issues page](https://github.com/ask-marcel/ask-marcel-office-cli/issues).

## License

MIT © Vincent Delacourt
