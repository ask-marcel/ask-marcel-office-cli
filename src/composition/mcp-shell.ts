/*
 * The MCP gateway — `ask-marcel-office mcp`, for whichever package it is given.
 *
 * Four gateway tools plus the package's own (the single package adds `login`
 * and `status`, see `auth-mcp.ts`), not one per command. The registry already carries everything a tool manifest
 * needs, but exposing one MCP tool per command would inject hundreds of KB of
 * schema into every client session — the exact token bloat this CLI exists to
 * avoid. So discovery is three hops (list -> docs -> run) and the tool list is
 * a constant. See docs/adr/0001-mcp-server-layer.md.
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
import { describeWrites, WRITE_EFFECTS, writeGroups } from '../use-cases/commands/command-effect.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { setDateZone } from '../use-cases/commands/date-zone.ts';
import { buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs-builders.ts';
import { CATEGORY_ORDER } from '../use-cases/commands/docs-render.ts';
import { resolveCommand } from '../use-cases/commands/resolve-command.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import { resolveDateZone } from './date-zone.ts';
import { unknownCommandMessage } from '../use-cases/commands/unknown-command.ts';
import { buildRenderContext, runRegistryCommand } from './run-registry-command.ts';

const MCP_REMEDY = 'Call list-commands to see every command.';

/*
 * The gateway is core (package split, step 13): it serves the package it is
 * given and imports no registry and no lifecycle command. The composition root
 * hands it the package (`mcp.ts` for the single package).
 */
type McpShellDeps = {
  readonly commandPackage: CommandPackage;
  readonly graph: ReadGraph & WriteGraph;
  readonly fs: FileSystem;
  readonly version?: string;
  /** Registers the package's own tools (`login`, `status`…) after the four gateway tools. */
  readonly addTools?: (server: McpServer) => void;
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

const buildMcpShell = (deps: McpShellDeps): McpServer => {
  // Same rule as the CLI: named days resolve in `ASKMARCEL_TZ`, else the machine's zone.
  setDateZone(resolveDateZone(undefined).zone);
  const { commandPackage, graph, fs } = deps;
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
  const readCommandNames = Object.entries(registry)
    .filter(([, c]) => c.meta.effect === 'read')
    .map(([n]) => n);
  const writeCommandNames = Object.entries(registry)
    .filter(([, c]) => c.meta.effect !== 'read')
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));
  const writesAreDestructive = writeGroups(registry).some((group) => WRITE_EFFECTS[group.effect].destructive);

  server.registerTool(
    'list-commands',
    {
      title: 'List commands',
      description: `Discover what this Microsoft 365 CLI can do. Returns the terse manifest — {name, summary, category} per command — for all ${readCommandNames.length + writeCommandNames.length} commands, or one category. START HERE, then use get-command-docs on the command you picked to learn its params. Categories: ${CATEGORY_LIST}.`,
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
      description:
        'Full Markdown docs for ONE command: every option, its Graph endpoint, an example, and the response shape. Call this after list-commands and before run-command — it tells you exactly which params to pass. Also covers the lifecycle commands (login/status/logout/docs/help-json/mcp).',
      inputSchema: {
        command: z.string().describe('Command name, e.g. `list-mail-messages`.'),
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
    params: z.record(z.string(), z.string()).optional().describe('Command options WITHOUT the `--` prefix, e.g. {"top":"10","folderId":"inbox"}. See get-command-docs.'),
    outputPath: z
      .string()
      .optional()
      .describe('For commands returning bytes/text: write the body to this file instead of inlining it. Avoids flooding context with a multi-MB base64 blob.'),
    outputDir: z.string().optional().describe('For the extract-*-images commands: write each image here instead of inlining it.'),
  };

  const runResolved = async (
    commandName: string,
    params: Record<string, string> | undefined,
    outputPath: string | undefined,
    outputDir: string | undefined,
    wantMutating: boolean
  ): Promise<CallToolResult> => {
    const resolved = resolveCommand(registry, commandName);
    if (!resolved.ok) return unknownCommand(resolved.error.name, resolved.error.available);
    const isMutating = resolved.value.command.meta.effect !== 'read';
    // Gate BEFORE executing: a write routed through run-command would silently
    // mutate under a `readOnlyHint: true` tool the client may have auto-approved.
    if (isMutating && !wantMutating)
      return failText(
        `"${resolved.value.name}" writes to Microsoft 365, so it is not available through run-command (which is declared read-only). Use run-write-command instead. Write commands: ${writeCommandNames.join(', ')}.`
      );
    if (!isMutating && wantMutating)
      return failText(`"${resolved.value.name}" is a read command — use run-command. run-write-command only accepts: ${writeCommandNames.join(', ')}.`);
    const request = { name: resolved.value.name, command: resolved.value.command, params: params ?? {}, outputPath, outputDir, surface: 'mcp' as const };
    const result = await runRegistryCommand({ graph, fs, commandPackage }, request);
    if (!result.ok)
      return failText(`${result.error.message}${signInPointer(result.error.code, commandPackage)}`, result.error.code, result.error.source, result.error.retryAfterSeconds);
    return okText(renderToString(result.value, 'text', buildRenderContext(resolved.value.name, resolved.value.command, 'mcp', params ?? {}), commandPackage));
  };

  server.registerTool(
    'run-command',
    {
      title: 'Run a read command',
      description: `Run any of the ${readCommandNames.length} READ commands (mail, files, calendar, chats, search, document conversion). Cannot write: the ${writeCommandNames.length} write commands are rejected here and live on run-write-command. Call get-command-docs first to learn the params.`,
      inputSchema: runToolInput,
      annotations: { readOnlyHint: true },
    },
    async ({ command, params, outputPath, outputDir }): Promise<CallToolResult> => runResolved(command, params, outputPath, outputDir, false)
  );

  server.registerTool(
    'run-write-command',
    {
      title: 'Run a write command',
      description: `Run one of the ${writeCommandNames.length} commands that WRITE to Microsoft 365. ${describeWrites(registry)}. This CLI has no other write surface. Read commands are rejected here; use run-command.`,
      inputSchema: runToolInput,
      // Not read-only. Destructive only when one of the write classes is (none
      // today: a draft is unsent, and a temporary upload is a copy of the user's
      // own attachment that the command tries to delete; nothing is overwritten
      // or transmitted).
      annotations: { readOnlyHint: false, destructiveHint: writesAreDestructive },
    },
    async ({ command, params, outputPath, outputDir }): Promise<CallToolResult> => runResolved(command, params, outputPath, outputDir, true)
  );

  // The package's own tools (login and status in the single package) sit beside the gateway tools.
  deps.addTools?.(server);

  return server;
};

export { buildMcpShell, errText, okText };
export type { McpShellDeps };
