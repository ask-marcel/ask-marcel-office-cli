/*
 * The tools of the auth package's MCP server: `login` and `status`
 * (docs/plans/2026-10-01-package-split.md). `auth-server.ts` gives them to its
 * own `McpServer`; the read and write gateways (on core's `mcp-shell.ts`) do
 * not have them. Like every MCP module, nothing here writes to stdout: a
 * tool returns its text.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import type { AuthLadder, LoginAuthFactory } from '../infra/auth.ts';
import { renderToString } from '@ask-marcel/office-core/presenter/render-to-string.ts';
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { fillNames, fillNamesIn } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import * as login from '../use-cases/commands/login.ts';
import { buildLoginSummary } from '../use-cases/commands/login-status.ts';
import * as status from '../use-cases/commands/status.ts';
import { errText, okText } from '@ask-marcel/office-core/composition/mcp-shell.ts';

type AuthToolsDeps = {
  readonly auth: AuthLadder;
  /**
   * Builds the AuthManager for an interactive `login` run (it may recapture
   * secondary tokens via the browser, unlike the command-path `auth`). Mirrors
   * `auth-cli.ts`: when omitted (tests), the login tool falls back to `auth`.
   */
  readonly makeLoginAuth?: LoginAuthFactory;
  /** The package the tools are part of: its bins fill the placeholders of their text. */
  readonly names: NameSource;
};

const addAuthTools = (server: McpServer, deps: AuthToolsDeps): void => {
  const { auth, names } = deps;
  const failText = (message: string, code?: string, source?: 'cli'): CallToolResult => errText(message, code, source, undefined, names);
  server.registerTool(
    'login',
    {
      title: 'Sign in to Microsoft 365',
      description: fillNames(
        "Authenticate via the Teams web client (cached token -> refresh -> browser). Opens a browser window on this machine when it needs one. Run this when a command fails with an auth error: the elevated (M365) token lapses roughly hourly and only a browser can recapture it. SLOW BY NATURE — a browser sign-in measured 37-64 seconds even with no MFA prompt, which straddles the 60s default MCP tool timeout, so this call may time out. If it does, the sign-in often completed anyway: re-run your original command before calling login again. Raise MCP_TOOL_TIMEOUT (or your client's equivalent) to ~300000 to avoid it. First-time setup, where an MFA prompt adds minutes, is better done in a terminal: `{auth-bin} login`.",
        names
      ),
      inputSchema: {
        force: z
          .boolean()
          .optional()
          .describe(
            'Ignore the cache and re-capture every token via the browser. A plain login already recaptures the elevated token when it is missing and leaves the browser session alone; force clears that session first, which deletes the 90-day "Stay signed in" cookie and usually forces the user to enter their password. Use it only when a tier is stuck after a plain login.'
          ),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    },
    async ({ force }): Promise<CallToolResult> => {
      const loginAuth = deps.makeLoginAuth ? deps.makeLoginAuth() : auth;
      const result = await login.execute(loginAuth, { force: force ?? false });
      if (!result.ok) return failText(result.error.type === 'auth_cancelled' ? 'Authentication cancelled' : result.error.message);
      const info = await loginAuth.getTokenInfo();
      if (!info.ok) return failText(info.error.message);
      const summary = buildLoginSummary({
        elevatedAvailable: info.value.elevated.available,
        chatsvcaggAvailable: info.value.chatsvcagg.available,
        ic3Available: info.value.ic3.available,
      });
      return okText(renderToString(fillNamesIn(summary, names), 'text'));
    }
  );

  server.registerTool(
    'status',
    {
      title: 'Show the cached tokens',
      description:
        'Show the four cached tokens (basic, elevated, chatsvcagg and ic3). For each token, it shows if the token is available and the seconds before its expiry. It also shows the scopes, the refresh method, and the data that the token lets you read. It reads only the token cache, and it does not make a Graph call or open a browser. Call it before login. If the token that a command uses is available, a sign-in is not necessary.',
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async (): Promise<CallToolResult> => {
      const result = await status.execute(auth);
      if (!result.ok) return failText(result.error.message, result.error.code, 'cli');
      return okText(renderToString(fillNamesIn(result.value, names), 'text'));
    }
  );
};

export { addAuthTools };
export type { AuthToolsDeps };
