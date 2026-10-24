import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';

/**
 * Lifecycle commands aren't backed by a Graph endpoint and aren't in the
 * `commands` registry — `docs` and `help-json` live in `cli-shell.ts`, and
 * `mcp` in `main.ts`. But an LLM that consumes `help-json` (the recommended
 * token-efficient surface) needs to see them too, otherwise it can't fetch
 * per-command docs. These stub entries surface them under category
 * 'lifecycle' with no graphMethod / options (the path template is a
 * `(meta) ...` description so the JSON shape stays uniform with regular
 * commands). The single package hands them to the shells as its
 * `CommandPackage.lifecycle` (`composition/single-package.ts`). `login`,
 * `logout` and `status` are the auth package's (`auth-lifecycle-entries.ts`),
 * and the single package does not list them.
 */
export const LIFECYCLE_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  {
    name: 'docs',
    summary:
      'Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`). Pass the command name as a POSITIONAL argument — there is no `--command` flag. For lifecycle commands (docs/help-json/mcp) prints the same --help that command would.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders Markdown from the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    positionalArguments: [
      {
        name: 'command',
        required: true,
        description:
          'Name of the command to show docs for (e.g. `list-mail-messages`). Run `{bin} --help` or `{bin} help-json` for the full list. Passed as a bare positional — do NOT prefix with `--command`.',
      },
    ],
    example: '{bin} docs list-mail-messages',
    responseShape: 'Markdown text on stdout (NOT JSON-wrapped — this is the one command whose stdout is plain text).',
  },
  {
    name: 'help-json',
    summary:
      'Print the machine-readable command manifest as JSON. For fresh-session discovery use `--terse --category <name>` (~8 KB for one category). The unflagged form is the *full* reference (every option / example / response shape per command) and is roughly 14× the size of `{bin} --help` — reach for it only after `--terse` has narrowed the search. `--terse` alone projects each entry to `{name, summary, category}` (summary compacted to its first sentence). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: '{bin} help-json',
    responseShape: '{ package, version, generatedAt, commands: [{ name, summary, category, ... }, ...] }',
  },
  {
    name: 'mcp',
    summary:
      'Start the Model Context Protocol server over stdio, exposing every command to an MCP client (Claude Code, Claude Desktop, Cursor, …) as three gateway tools (list-commands, get-command-docs, run-command). Runs in the foreground and serves until the client disconnects — a server entry point, not a one-shot command, so a shell- or MCP-driving agent never calls it directly. Register with `claude mcp add --transport stdio --scope user {bin} -- {bin} mcp`. See docs/USAGE.md (MCP server).',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) serves the command registry over MCP stdio; not a Graph endpoint',
    graphDocsUrl: 'https://modelcontextprotocol.io/',
    options: [],
    example: '{bin} mcp',
    responseShape: 'None — a long-lived stdio server emitting JSON-RPC frames until the client disconnects (never a JSON envelope).',
  },
];
