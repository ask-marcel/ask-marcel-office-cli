<div align="center">

# ask-marcel-office

### The Microsoft 365 command line built for AI agents

**One sign-in. No Azure app registration. Every document as clean markdown.**

[![npm version](https://img.shields.io/npm/v/ask-marcel-office-cli.svg?logo=npm&color=cb3837)](https://www.npmjs.com/package/ask-marcel-office-cli)
[![license: MIT](https://img.shields.io/npm/l/ask-marcel-office-cli.svg?color=blue)](LICENSE)
[![types included](https://img.shields.io/npm/types/ask-marcel-office-cli?logo=typescript&color=3178c6)](docs/USAGE.md)

Outlook · OneDrive · SharePoint · Calendar · Excel · Teams · Planner · To Do · OneNote · People

[Install](#install-in-60-seconds) · [See it work](#see-it-work) · [What it reaches](#what-your-agent-can-reach) · [Files to markdown](#any-file-becomes-markdown) · [Library](#embed-it-as-a-typescript-library) · [All 211 commands](docs/COMMANDS.md)

</div>

---

Your agent is smart enough to answer *"what did Contoso say about the Q3 budget, and what's in the attached deck?"* It just can't see your mailbox.

`ask-marcel-office` gives any tool-calling LLM (Claude Code, Cursor, Cline, an MCP server, your own loop) eyes on your entire Microsoft 365:

```bash
npm i -g ask-marcel-office-cli
ask-marcel-office login          # your normal Microsoft sign-in, in a browser, once
ask-marcel-office list-mail-messages --top 5
```

That is the whole setup. **No Azure app registration. No tenant-admin consent. No client secrets.** And little for a runaway agent to break: the command surface is read-mostly by design. A write leaves an unsent draft, or a temporary OneDrive file that the command tries to delete.

<p align="center">
  <img src="https://raw.githubusercontent.com/ask-marcel/ask-marcel-office-cli/main/docs/demo.gif" alt="ask-marcel-office converting an Outlook .msg file into clean markdown in one command, offline" width="880">
</p>

The product pages are now in the [read package README](read/README.md): the comparison with raw Microsoft Graph, the three walls, "See it work", what your agent can reach, files to markdown, the agent loop, MCP and the library.

## Install in 60 seconds

```bash
# Bun >=1.0 or Node >=20 · macOS, Windows, Linux
npm i -g ask-marcel-office-cli

ask-marcel-office login             # browser opens once, tokens cached
ask-marcel-office my-quick-context  # who am I + my IDs, one round trip
ask-marcel-office list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel-office search-onedrive-files --drive-id "b!abc..." --query "Q3 budget"
```

## Deep docs

- **[All 211 commands](docs/COMMANDS.md)**: per-category tables with required params + Graph endpoint
- **[Usage guide](docs/USAGE.md)**: output formats, OData passthrough, `--output-path`, pagination, library API, architecture, configuration
- **[Machine-readable manifest](docs/commands.json)**: JSON for programmatic discovery, also importable via `import manifest from 'ask-marcel-office-cli/commands.json'`
- **[QA playbook](docs/QA-PLAYBOOK.md)**: the repeatable full-surface health check run before each release

## Roadmap

Read-only stays the default forever. Coverage grows out of real LLM workflows as they come up: suggestions, requests, and pull requests welcome on the [issues page](https://github.com/ask-marcel/ask-marcel-office-cli/issues).

## Built with

- **Bun + TypeScript**: single-package install, Node >=20 fallback, `Result<T, E>` at every IO boundary, a 100% coverage gate on every tier
- **Microsoft Graph v1.0**: the public API surface, no beta endpoints in production code
- **Playwright**: the headed browser behind the one-time sign-in

## License

MIT © Vincent Delacourt

---

<div align="center">

**If this saved your agent a few round trips, a ⭐ helps other agents' humans find it.**

</div>
