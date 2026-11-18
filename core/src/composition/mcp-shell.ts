/*
 * The MCP gateway — `ask-marcel-office mcp`, for whichever package it is given.
 *
 * At most five gateway tools plus any tools the package adds (`addTools`), not
 * one per command: `run-command` only when the package has a read command,
 * `run-write-command` only when it has a write that does not delete, and
 * `run-destructive-command` only when it has a command that deletes (the one
 * tool with the destructive mark, mcp-run-tools.ts). The registry already
 * carries everything a tool manifest needs, but exposing one MCP tool per
 * command would inject hundreds of KB of schema into every client session — the
 * exact token bloat this CLI exists to avoid. So discovery is three hops (list
 * -> docs -> run) and the tool list of a package is a constant. See
 * docs/adr/0001-mcp-server-layer.md.
 *
 * STDOUT IS THE JSON-RPC CHANNEL. Nothing in this module may write to it —
 * one stray `process.stdout.write` corrupts the frame stream and the client
 * drops the connection. That is why results go through `renderToString`
 * (pure) rather than the CLI's `render` (writes to stdout). The winston logger
 * is stderr-only, so it stays safe.
 *
 * Note that `InMemoryTransport` (what mcp.test.ts and mcp-shell.test.ts drive) does not touch
 * stdout at all, so a stray write is INVISIBLE to the unit tests. The bundle
 * smoke test in scripts/qa-bundle-smoke.ts is the gate that catches it.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import type { ReadGraph } from '../infra/read-graph.ts';
import type { WriteGraph } from '../infra/write-graph.ts';
import { renderErrorToString, renderToString } from '../presenter/render-to-string.ts';
import type { NameSource } from '../use-cases/commands/bin-names.ts';
import { fillNames } from '../use-cases/commands/bin-names.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import { setDateZone } from '../use-cases/commands/date-zone.ts';
import { buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs-builders.ts';
import { CATEGORY_ORDER } from '../use-cases/commands/docs-render.ts';
import type { RawParamValue } from '../use-cases/commands/param-keys.ts';
import { canonicalParams } from '../use-cases/commands/param-keys.ts';
import { resolveCommand } from '../use-cases/commands/resolve-command.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { LocalFileReader } from '../use-cases/commands/read-local-file.ts';
import type { RunRecorder } from '../use-cases/commands/record-run.ts';
import { resolveDateZone } from './date-zone.ts';
import { unknownCommandMessage } from '../use-cases/commands/unknown-command.ts';
import type { RunTool } from './mcp-run-tools.ts';
import { deleteToolDescription, runToolNames, runToolOf, runToolsPhrase, writesElsewhere, writeToolDescription, wrongToolMessage } from './mcp-run-tools.ts';
import { buildRenderContext, runRegistryCommand } from './run-registry-command.ts';

const MCP_REMEDY = 'Call list-commands to see every command.';

/*
 * The gateway is core (package split, step 13): it serves the package it is
 * given and imports no registry and no lifecycle command. The composition root
 * hands it the package (`mcp.ts` for the single package). `G` is the graph the
 * package's commands run on: by default the full client.
 */
type McpShellDeps<G = ReadGraph & WriteGraph> = {
  readonly commandPackage: CommandPackage<Command<G>>;
  readonly graph: G;
  readonly fs: FileSystem;
  readonly version?: string;
  /** Registers the package's own tools (`login`, `status`…) after the gateway tools. */
  readonly addTools?: (server: McpServer) => void;
  /** The run journal of the write package: with it, a write with a plan gets a run id (run-registry-command.ts). */
  readonly runs?: RunRecorder;
  /** The reader of a local file that a param of a write command names (run-registry-command.ts). */
  readonly files?: LocalFileReader;
};

// The SDK's own result type rather than a hand-rolled one — it requires a
// MUTABLE `content` array, so a `readonly` shape here would typecheck against
// nothing and drift on the next SDK bump.
const okText = (text: string): CallToolResult => ({ content: [{ type: 'text', text }] });

// Failures come back as `isError: true` rather than a thrown exception: the
// spec routes tool errors to the MODEL (which can correct itself) instead of
// to the client as a protocol fault. The body is the same text envelope the
// CLI prints, so `error-hints.ts` remedies reach MCP callers unchanged; `names`
// fills their bin placeholders.
const errText = (message: string, code?: string, source?: Parameters<typeof renderErrorToString>[3], retryAfterSeconds?: number, names?: NameSource): CallToolResult => ({
  content: [{ type: 'text', text: renderErrorToString(message, 'text', code, source, retryAfterSeconds, names) }],
  isError: true,
});

// The failures a sign-in cures. On a server without the `login` tool (any but
// the auth package's), the message also names the server that has it.
const SIGN_IN_CODES: ReadonlySet<string> = new Set(['not_authenticated', 'secondary_token_unavailable', 'InvalidAuthenticationToken', 'TokenExpired']);

const signInPointer = (code: string | undefined, names: NameSource): string =>
  code !== undefined && SIGN_IN_CODES.has(code) && names.bins.auth !== names.binName
    ? ` From an MCP client, call the \`login\` tool of the ${names.bins.auth} server (\`${names.bins.auth} mcp\`).`
    : '';

// Derived from the single curated order in docs-render, never hardcoded, so a new category can never
// drift out of sync with what the category filter actually accepts (docs-render.test pins completeness).
const CATEGORY_LIST = CATEGORY_ORDER.join(', ');

const buildMcpShell = <G>(deps: McpShellDeps<G>): McpServer => {
  // Same rule as the CLI: named days resolve in `ASKMARCEL_TZ`, else the machine's zone.
  setDateZone(resolveDateZone(undefined).zone);
  const { commandPackage, graph, fs, runs, files } = deps;
  const { registry } = commandPackage;
  const version = deps.version ?? '0.0.0';
  const server = new McpServer({ name: commandPackage.binName, version });
  // The gateway's own text (manifest, docs pages, hints) names commands through
  // the bin placeholders; this package's names fill them (bin-names.ts).
  const fill = (text: string): string => fillNames(text, commandPackage);
  const failText = (message: string, code?: string, source?: Parameters<typeof renderErrorToString>[3], retryAfterSeconds?: number): CallToolResult =>
    errText(message, code, source, retryAfterSeconds, commandPackage);
  const unknownCommand = (name: string, available: ReadonlyArray<string>): CallToolResult =>
    failText(unknownCommandMessage(name, available, MCP_REMEDY, commandPackage, 'mcp'), 'cli_unknown_command');

  // Derived, never hardcoded — and that goes for prose about it too. This comment
  // used to name the split as a pair of literals, and those literals were stale
  // within two releases: exactly the failure the derivation below exists to
  // prevent. Read the counts off the package's registry, never off a sentence.
  // The same holds for what a write does: each command's `effect` decides its tool,
  // and the wording and annotations come from the effect table in command-effect.ts.
  const toolNames = runToolNames(registry);
  const readCommandNames = toolNames['run-command'];
  // Each run tool exists only when the package has a command of its kind, so a
  // server never offers a tool that can only refuse. The other kind is on the
  // server of another package.
  const hasReads = readCommandNames.length > 0;
  const runTools = runToolsPhrase(toolNames);
  const lifecycleNames = commandPackage.lifecycle.map((entry) => entry.name).join('/');
  const otherServer = (role: 'read' | 'write'): string => `another MCP server (\`${commandPackage.bins[role]} mcp\`)`;

  server.registerTool(
    'list-commands',
    {
      title: 'List commands',
      description: `Discover what this Microsoft 365 CLI can do. Returns the terse manifest — {name, summary, category} per command — for all ${readCommandNames.length + toolNames.writes.length} commands, or one category. START HERE, then use get-command-docs on the command you picked to learn its params. Categories: ${CATEGORY_LIST}.`,
      inputSchema: {
        category: z.string().optional().describe(`Filter to one category (${CATEGORY_LIST}). Omit for every command — larger, but one round-trip.`),
      },
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async ({ category }): Promise<CallToolResult> => {
      // The manifest is the gateway's own text only, so its render is filled whole.
      const manifest = buildTerseManifest(commandPackage, version);
      if (category === undefined) return okText(fill(renderToString(manifest, 'text')));
      const filtered = filterManifestByCategory(manifest, category);
      if (!filtered.ok) return failText(`Unknown category "${filtered.error.category}". Available categories: ${filtered.error.available.join(', ')}.`);
      return okText(fill(renderToString(filtered.value, 'text')));
    }
  );

  server.registerTool(
    'get-command-docs',
    {
      title: 'Get command docs',
      description: `Full Markdown docs for ONE command: every option, its Graph endpoint, an example, and the response shape. Call this after list-commands and before ${runTools} — it tells you exactly which params to pass. Also covers the lifecycle commands (${lifecycleNames}).`,
      inputSchema: {
        command: z.string().describe(`Command name, e.g. \`${commandPackage.exampleCommand}\`.`),
      },
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async ({ command }): Promise<CallToolResult> => {
      // When resolution fails, pass the raw name through so lifecycle entries
      // (which are not in the registry) still render and unknown names still
      // get renderSingleCommand's own `available` list.
      const resolved = resolveCommand(registry, command);
      const result = renderSingleCommand(commandPackage, resolved.ok ? resolved.value.name : command);
      if (!result.ok) return unknownCommand(result.error.name, result.error.available);
      return okText(fill(result.value));
    }
  );

  // The read/write split exists so the read commands keep an honest
  // `readOnlyHint: true` (which lets a client auto-approve them) instead of
  // losing it to the handful of write commands sharing one tool.
  const runToolInput = {
    command: z.string().describe('Command name from list-commands.'),
    params: z
      .record(z.string(), z.union([z.string(), z.number(), z.boolean()]))
      .optional()
      .describe(
        'Command options as get-command-docs shows them, e.g. {"message-id":"<id>","top":5}. A key is the flag name, with or without its `--` prefix, or its camelCase form (`messageId`). A value is a string, a number or a boolean.'
      ),
    outputPath: z
      .string()
      .optional()
      .describe('For commands returning bytes/text: write the body to this file instead of inlining it. Avoids flooding context with a multi-MB base64 blob.'),
    outputDir: z
      .string()
      .optional()
      .describe('For commands that return images (the extract-*-images commands, a zip converted with `include-images`): write each image here instead of inlining it.'),
  };

  // What every run needs; the run journal only when the package has one.
  const runDeps = { graph, fs, commandPackage, ...(runs === undefined ? {} : { runs }), ...(files === undefined ? {} : { files }) };

  // The write tool also takes `dryRun`: a write with a plan then shows its changes and makes none.
  const writeToolInput = {
    ...runToolInput,
    dryRun: z
      .boolean()
      .optional()
      .describe(
        'true: show each change that the command would make (the target and the ids before it), and make none. Only the commands whose docs show "Dry-run and undo" take it.'
      ),
  };

  const runResolved = async (
    commandName: string,
    rawParams: Readonly<Record<string, RawParamValue>> | undefined,
    outputPath: string | undefined,
    outputDir: string | undefined,
    called: RunTool,
    dryRun: boolean
  ): Promise<CallToolResult> => {
    const resolved = resolveCommand(registry, commandName);
    if (!resolved.ok) return unknownCommand(resolved.error.name, resolved.error.available);
    // Gate BEFORE executing: each command runs only on the tool whose marks fit its effect (mcp-run-tools.ts).
    const wrongTool = wrongToolMessage(resolved.value.name, runToolOf(resolved.value.command.meta.effect), called, toolNames);
    if (wrongTool !== undefined) return failText(wrongTool);
    // The keys arrive as the docs name them (`message-id`); a command reads `messageId` (param-keys.ts).
    const canonical = canonicalParams(rawParams ?? {});
    if (!canonical.ok)
      return failText(`\`${canonical.error.first}\` and \`${canonical.error.second}\` are the same parameter. Give it one time.`, canonical.error.type, 'validation');
    const params = canonical.value;
    const request = { name: resolved.value.name, command: resolved.value.command, params, outputPath, outputDir, surface: 'mcp' as const, dryRun };
    const result = await runRegistryCommand(runDeps, request);
    if (!result.ok)
      return failText(`${result.error.message}${signInPointer(result.error.code, commandPackage)}`, result.error.code, result.error.source, result.error.retryAfterSeconds);
    return okText(renderToString(result.value, 'text', buildRenderContext(resolved.value.name, resolved.value.command, 'mcp', params), commandPackage));
  };

  if (hasReads)
    server.registerTool(
      'run-command',
      {
        title: 'Run a read command',
        description: `Run any of the ${readCommandNames.length} READ commands (mail, files, calendar, chats, search, document conversion). Cannot write: ${writesElsewhere(toolNames, otherServer('write'))}. Call get-command-docs first to learn the params.`,
        inputSchema: runToolInput,
        annotations: { readOnlyHint: true },
      },
      async ({ command, params, outputPath, outputDir }): Promise<CallToolResult> => runResolved(command, params, outputPath, outputDir, 'run-command', false)
    );

  const readsElsewhere = hasReads ? 'Read commands are rejected here; use run-command.' : `The read commands are on ${otherServer('read')}.`;
  // A write tool, registered only when the package has a command for it.
  const writeTool = (tool: RunTool, title: string, description: string, destructive: boolean): void => {
    if (toolNames[tool].length === 0) return;
    server.registerTool(tool, { title, description, inputSchema: writeToolInput, annotations: { readOnlyHint: false, destructiveHint: destructive } }, async (input) =>
      runResolved(input.command, input.params, input.outputPath, input.outputDir, tool, input.dryRun === true)
    );
  };
  // Not read-only, and not destructive: a draft is unsent, a temporary upload is
  // a copy of the user's own attachment that the command tries to delete, and a
  // state change has an undo. The commands that delete are on
  // run-destructive-command, so this mark never covers a delete.
  writeTool('run-write-command', 'Run a write command', writeToolDescription(registry, toolNames, readsElsewhere), false);
  writeTool('run-destructive-command', 'Run a command that deletes', deleteToolDescription(registry, toolNames, hasReads ? 'run-command' : otherServer('read')), true);

  // The tools that the package adds, if any, sit beside the gateway tools.
  deps.addTools?.(server);

  return server;
};

export { buildMcpShell, errText, okText };
export type { McpShellDeps };
