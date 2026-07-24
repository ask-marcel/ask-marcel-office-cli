/*
 * The MCP gateway — `ask-marcel-office mcp`.
 *
 * Five tools, not 184. The registry already carries everything a tool manifest
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
 * Note that `InMemoryTransport` (what mcp.test.ts drives) does not touch
 * stdout at all, so a stray write is INVISIBLE to the unit tests. The bundle
 * smoke test in scripts/qa-bundle-smoke.ts is the gate that catches it.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import type { AuthManager } from '../infra/auth.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { renderErrorToString, renderToString } from '../presenter/render-to-string.ts';
import { buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs.ts';
import { CATEGORY_ORDER } from '../use-cases/commands/docs-render.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import * as login from '../use-cases/commands/login.ts';
import { buildLoginSummary } from '../use-cases/commands/login-status.ts';
import { resolveCommand } from '../use-cases/commands/resolve-command.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { LoginAuthFactory } from './build-deps.ts';
import { buildSizeHintContext, runRegistryCommand } from './run-registry-command.ts';

const PACKAGE_NAME = 'ask-marcel-office-cli';

type BuildMcpServerDeps = {
  readonly auth: AuthManager;
  readonly graph: GraphClient;
  readonly fs: FileSystem;
  readonly version?: string;
  /**
   * Builds the AuthManager for an interactive `login` run (it may recapture
   * secondary tokens via the browser, unlike the command-path `auth`). Mirrors
   * `cli.ts`: when omitted (tests), the login tool falls back to `auth`.
   */
  readonly makeLoginAuth?: LoginAuthFactory;
};

// The SDK's own result type rather than a hand-rolled one — it requires a
// MUTABLE `content` array, so a `readonly` shape here would typecheck against
// nothing and drift on the next SDK bump.
const okText = (text: string): CallToolResult => ({ content: [{ type: 'text', text }] });

// Failures come back as `isError: true` rather than a thrown exception: the
// spec routes tool errors to the MODEL (which can correct itself) instead of
// to the client as a protocol fault. The body is the same text envelope the
// CLI prints, so `error-hints.ts` remedies reach MCP callers unchanged.
const errText = (message: string, code?: string, source?: Parameters<typeof renderErrorToString>[3], retryAfterSeconds?: number): CallToolResult => ({
  content: [{ type: 'text', text: renderErrorToString(message, 'text', code, source, retryAfterSeconds) }],
  isError: true,
});

// Derived, never hardcoded. The ADR said 182/3; the registry says 180/4. Any
// literal here would be stale the next time a command lands.
const readCommandNames = Object.entries(cmdRegistry)
  .filter(([, c]) => c.meta.mutates !== true)
  .map(([n]) => n);
const writeCommandNames = Object.entries(cmdRegistry)
  .filter(([, c]) => c.meta.mutates === true)
  .map(([n]) => n)
  .toSorted((a, b) => a.localeCompare(b));

// Derived from the single curated order in docs-render, never hardcoded, so a new category can never
// drift out of sync with what the category filter actually accepts (docs-render.test pins completeness).
const CATEGORY_LIST = CATEGORY_ORDER.join(', ');

const buildMcpServer = (deps: BuildMcpServerDeps): McpServer => {
  const { auth, graph, fs } = deps;
  const version = deps.version ?? '0.0.0';
  const server = new McpServer({ name: 'ask-marcel-office', version });

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
      const manifest = buildTerseManifest(cmdRegistry, PACKAGE_NAME, version);
      if (category === undefined) return okText(renderToString(manifest, 'text'));
      const filtered = filterManifestByCategory(manifest, category);
      if (!filtered.ok) return errText(`Unknown category "${filtered.error.category}". Available categories: ${filtered.error.available.join(', ')}.`);
      return okText(renderToString(filtered.value, 'text'));
    }
  );

  server.registerTool(
    'get-command-docs',
    {
      title: 'Get command docs',
      description:
        'Full Markdown docs for ONE command: every option, its Graph endpoint, an example, and the response shape. Call this after list-commands and before run-command — it tells you exactly which params to pass. Also covers the lifecycle commands (login/logout/update/docs/help-json/mcp).',
      inputSchema: {
        command: z.string().describe('Command name, e.g. `list-mail-messages`.'),
      },
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async ({ command }): Promise<CallToolResult> => {
      // When resolution fails, pass the raw name through so lifecycle entries
      // (which are not in the registry) still render and unknown names still
      // get renderSingleCommand's own `available` list.
      const resolved = resolveCommand(cmdRegistry, command);
      const result = renderSingleCommand(cmdRegistry, resolved.ok ? resolved.value.name : command);
      if (!result.ok) return errText(`Unknown command "${result.error.name}". Call list-commands to see every command.`, 'cli_unknown_command');
      return okText(result.value);
    }
  );

  // The read/write split exists so the 180 read commands keep an honest
  // `readOnlyHint: true` (which lets a client auto-approve them) instead of
  // losing it to the 4 draft commands sharing one tool.
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
    const resolved = resolveCommand(cmdRegistry, commandName);
    if (!resolved.ok) return errText(`Unknown command "${resolved.error.name}". Call list-commands to see every command.`, 'cli_unknown_command');
    const isMutating = resolved.value.command.meta.mutates === true;
    // Gate BEFORE executing: a write routed through run-command would silently
    // mutate under a `readOnlyHint: true` tool the client may have auto-approved.
    if (isMutating && !wantMutating)
      return errText(
        `"${resolved.value.name}" writes to Microsoft 365, so it is not available through run-command (which is declared read-only). Use run-write-command instead. Write commands: ${writeCommandNames.join(', ')}.`
      );
    if (!isMutating && wantMutating)
      return errText(`"${resolved.value.name}" is a read command — use run-command. run-write-command only accepts: ${writeCommandNames.join(', ')}.`);
    const request = { name: resolved.value.name, command: resolved.value.command, params: params ?? {}, outputPath, outputDir, surface: 'mcp' as const };
    const result = await runRegistryCommand({ graph, fs }, request);
    if (!result.ok) return errText(result.error.message, result.error.code, result.error.source, result.error.retryAfterSeconds);
    return okText(renderToString(result.value, 'text', buildSizeHintContext(resolved.value.name, resolved.value.command, 'mcp')));
  };

  server.registerTool(
    'run-command',
    {
      title: 'Run a read command',
      description: `Run any of the ${readCommandNames.length} READ commands (mail, files, calendar, chats, search, document conversion). Cannot write: the ${writeCommandNames.length} draft commands are rejected here and live on run-write-command. Call get-command-docs first to learn the params.`,
      inputSchema: runToolInput,
      annotations: { readOnlyHint: true },
    },
    async ({ command, params, outputPath, outputDir }): Promise<CallToolResult> => runResolved(command, params, outputPath, outputDir, false)
  );

  server.registerTool(
    'run-write-command',
    {
      title: 'Run a write command',
      description: `Run one of the ${writeCommandNames.length} commands that WRITE to Microsoft 365: ${writeCommandNames.join(', ')}. Each only creates or updates an UNSENT mail draft — this CLI cannot send mail, and has no other write surface. Read commands are rejected here; use run-command.`,
      inputSchema: runToolInput,
      // Not read-only, but not destructive either: every write here produces an
      // unsent draft. Nothing is overwritten or transmitted.
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ command, params, outputPath, outputDir }): Promise<CallToolResult> => runResolved(command, params, outputPath, outputDir, true)
  );

  server.registerTool(
    'login',
    {
      title: 'Sign in to Microsoft 365',
      description:
        "Authenticate via the Teams web client (cached token -> refresh -> browser). Opens a browser window on this machine when it needs one. Run this when a command fails with an auth error: the elevated (M365) token lapses roughly hourly and only a browser can recapture it. SLOW BY NATURE — a browser sign-in measured 37-64 seconds even with no MFA prompt, which straddles the 60s default MCP tool timeout, so this call may time out. If it does, the sign-in often completed anyway: re-run your original command before calling login again. Raise MCP_TOOL_TIMEOUT (or your client's equivalent) to ~300000 to avoid it. First-time setup, where an MFA prompt adds minutes, is better done in a terminal: `ask-marcel-office login`.",
      inputSchema: {
        force: z
          .boolean()
          .optional()
          .describe(
            'Ignore the cache and re-capture every token via the browser. A plain login already recaptures the elevated token when it is missing; use force to refresh every tier unconditionally.'
          ),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    },
    async ({ force }): Promise<CallToolResult> => {
      const loginAuth = deps.makeLoginAuth ? deps.makeLoginAuth() : auth;
      const result = await login.execute(loginAuth, { force: force ?? false });
      if (!result.ok) return errText(result.error.type === 'auth_cancelled' ? 'Authentication cancelled' : result.error.message);
      const info = await graph.getCachedTokenInfo();
      if (!info.ok) return errText(info.error.message);
      return okText(
        renderToString(
          buildLoginSummary({
            elevatedAvailable: info.value.elevated.available,
            chatsvcaggAvailable: info.value.chatsvcagg.available,
            ic3Available: info.value.ic3.available,
          }),
          'text'
        )
      );
    }
  );

  return server;
};

export { buildMcpServer };
export type { BuildMcpServerDeps };
