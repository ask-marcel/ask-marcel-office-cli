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
      'Start the Model Context Protocol server of the write package over stdio. It gives an MCP client (Claude Code, Claude Desktop, Cursor and others) the write commands through three gateway tools (list-commands, get-command-docs, run-write-command), and the undo tools undo and undo-apply. It runs in the foreground until the client disconnects. It is a server entry point, not a one-shot command, so an agent never calls it directly. Register it with `claude mcp add --transport stdio --scope user {bin} -- {bin} mcp`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) serves the write commands over MCP stdio; not a Graph endpoint',
    graphDocsUrl: 'https://modelcontextprotocol.io/',
    options: [],
    example: '{bin} mcp',
    responseShape: 'None: a long-lived stdio server that writes JSON-RPC frames until the client disconnects (never a JSON envelope).',
  },
  {
    name: 'undo',
    summary:
      'Show the undo of a write run, or do it with --apply. Give the run id as a POSITIONAL argument: the `runId` field of the output of a draft command that was not a dry-run. The undo of a new draft moves it to Deleted Items, and the undo of an update writes back the old fields. A change whose message is no longer a draft, or changed after the run, is passed over and reported. A run that is fully undone cannot be undone again, and the run journal keeps a run 30 days. Over MCP, the tools are `undo` (read-only) and `undo-apply` (destructive).',
    category: 'lifecycle',
    graphMethod: 'DELETE',
    graphPathTemplate:
      '(lifecycle) reads the run journal in ~/.ask-marcel/runs/, then for each change GET /me/messages/{id}, and with --apply DELETE /me/messages/{id} (a new draft) or PATCH /me/messages/{id} (an update)',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-delete',
    options: [{ name: 'apply', key: 'apply', required: false, description: 'Do the undo. Without it, the command only shows what the undo would do, and changes nothing.' }],
    positionalArguments: [{ name: 'runId', required: true, description: 'The runId field of the output of the write to reverse, for example 20261118T090507Z-0a1b2c3d.' }],
    example: 'undo 20261118T090507Z-0a1b2c3d --apply',
    responseShape:
      '`{ runId, command, startedAt, applied, changes: [{ action, target, after, undo, status, reason? }], warning? }`. `status` is `will-undo`, `undone`, `already-undone`, `changed`, `missing` or `failed`. A refusal has the error code `invalid_run_id`, `run_not_found`, `run_expired`, `run_unreadable`, `run_already_undone` or `journal_failed`.',
  },
];
