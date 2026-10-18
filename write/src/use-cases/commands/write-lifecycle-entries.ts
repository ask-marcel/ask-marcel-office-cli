import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';

/*
 * The lifecycle commands of the write package `@ask-marcel/office-write`: its
 * own `docs`, `help-json` and `mcp`. They are not in the write registry
 * (`write-commands.ts`), so these entries let `help-json` and `docs` show them.
 * The write package has no sign-in command: the auth package signs in. Its
 * MCP server has three gateway tools, and only `run-write-command` runs a
 * command.
 */
export const WRITE_LIFECYCLE_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  {
    name: 'docs',
    summary:
      'Print Markdown docs for one command of the write package. Give the command name as a POSITIONAL argument: there is no `--command` flag. For `docs`, `help-json` and `mcp`, the page is the same as the --help of that command.',
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
          'Name of the command to show docs for (for example `create-mail-draft`). Run `{bin} --help` or `{bin} help-json` for the full list. Give it as a bare positional argument: do NOT put `--command` before it.',
      },
    ],
    example: '{bin} docs create-mail-draft',
    responseShape: 'Markdown text on stdout (NOT JSON-wrapped: this is the one command whose stdout is plain text).',
  },
  {
    name: 'help-json',
    summary:
      'Print the machine-readable command manifest of the write package as JSON, with every option, example and response shape. `--terse` gives each entry as `{name, summary, category}` (the summary cut to its first sentence). Categories: lifecycle, mail, calendar.',
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
      'Start the Model Context Protocol server of the write package over stdio. It gives an MCP client (Claude Code, Claude Desktop, Cursor and others) the write commands through three gateway tools (list-commands, get-command-docs, run-write-command). It runs in the foreground until the client disconnects. It is a server entry point, not a one-shot command, so an agent never calls it directly. Register it with `claude mcp add --transport stdio --scope user {bin} -- {bin} mcp`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) serves the write commands over MCP stdio; not a Graph endpoint',
    graphDocsUrl: 'https://modelcontextprotocol.io/',
    options: [],
    example: '{bin} mcp',
    responseShape: 'None: a long-lived stdio server that writes JSON-RPC frames until the client disconnects (never a JSON envelope).',
  },
];
