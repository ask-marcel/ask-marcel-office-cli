// The auth cases of the MCP suite: the `login` and `status` tools, and the lifecycle docs
// of `login`. Moved out of mcp.test.ts with their text unchanged (package split, step 16),
// so the auth package takes this file whole; the helpers are copies of that file's.
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { buildMcpServer } from './mcp.ts';
import type { BuildMcpServerDeps } from './mcp.ts';

/**
 * Drives the REAL MCP protocol: a real Client and a real McpServer joined by
 * the SDK's in-memory transport pair. Only the secondary ports (Graph, auth,
 * filesystem) are fakes. So these tests exercise actual tool registration,
 * actual JSON-Schema input validation, and actual call dispatch — not our
 * assumption of them.
 */
const connect = async (overrides: Partial<BuildMcpServerDeps> = {}): Promise<Client> => {
  const server = buildMcpServer({
    auth: fakeAuthManager(),
    graph: fakeGraphClient(),
    fs: createFileSystemFake(),
    version: '9.9.9',
    ...overrides,
  });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => {
  const content = (result as { content: ReadonlyArray<{ type: string; text?: string }> }).content;
  return content.map((c) => c.text ?? '').join('');
};
const isError = (result: unknown): boolean => (result as { isError?: boolean }).isError === true;

describe('the MCP gateway an agent connects to', () => {
  it('marks the status tool read-only and idempotent, since it only reads the token cache', async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    const annotations = tools.find((t) => t.name === 'status')?.annotations ?? {};
    expect(annotations['readOnlyHint']).toBe(true);
    expect(annotations['idempotentHint']).toBe(true);
  });
});

describe('discovering what the CLI can do over MCP', () => {
  it('returns docs for a lifecycle command, which lives outside the registry', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'get-command-docs', arguments: { command: 'login' } }));
    expect(text).toContain('login');
  });
});

describe('authenticating from an MCP client', () => {
  it('reports which token tiers are available after a successful sign-in', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'login', arguments: {} });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('authenticated');
  });

  it('uses the login-configured auth manager, which may open a browser, rather than the fail-fast command-path one', async () => {
    let loginAuthUsed = false;
    const client = await connect({
      makeLoginAuth: () => {
        loginAuthUsed = true;
        return fakeAuthManager();
      },
    });
    await client.callTool({ name: 'login', arguments: {} });
    expect(loginAuthUsed).toBe(true);
  });

  it('passes force through so an agent can refresh the hourly-lapsing elevated token', async () => {
    let forced: boolean | undefined;
    const client = await connect({
      makeLoginAuth: () =>
        fakeAuthManager({
          getAccessToken: async (options?: { force?: boolean }) => {
            forced = options?.force;
            return ok('token' as never);
          },
        }),
    });
    await client.callTool({ name: 'login', arguments: { force: true } });
    expect(forced).toBe(true);
  });

  it('surfaces a cancelled sign-in as a tool error rather than a silent success', async () => {
    const client = await connect({
      makeLoginAuth: () => fakeAuthManager({ getAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } }) as never }),
    });
    const result = await client.callTool({ name: 'login', arguments: {} });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('cancelled');
  });

  it('surfaces a token-status read failure rather than reporting a sign-in that cannot be confirmed', async () => {
    const client = await connect({
      auth: fakeAuthManager({ getTokenInfo: async () => ({ ok: false, error: { type: 'auth_failed', message: 'cache unreadable' } }) }),
    });
    const result = await client.callTool({ name: 'login', arguments: {} });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('cache unreadable');
  });

  it('reports the four cached tokens through the status tool, with the refresh hint', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'status', arguments: {} });
    expect(isError(result)).toBe(false);
    for (const token of ['basic', 'elevated', 'chatsvcagg', 'ic3']) expect(textOf(result)).toContain(token);
    expect(textOf(result)).toContain('ask-marcel-office login');
  });

  // The text envelope has no error code line, so the message and the source are what the client sees.
  it('answers status with a tool error, from the cli source, that tells the client to sign in when the token cache has no session', async () => {
    const message = 'Not signed in: there is no cached token to inspect. Run `ask-marcel-office login`.';
    const client = await connect({ auth: fakeAuthManager({ getTokenInfo: async () => ({ ok: false, error: { type: 'auth_failed', message, code: 'not_authenticated' } }) }) });
    const result = await client.callTool({ name: 'status', arguments: {} });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toBe(`error: ${message}\ncode: not_authenticated\nsource: cli\n`);
  });

  it('reports the tokens that the sign-in just captured, read from the login-configured manager and not from the command-path one', async () => {
    const everyTier = { available: true, expiresInSeconds: 3600, scopes: [] };
    const client = await connect({
      makeLoginAuth: () =>
        fakeAuthManager({
          getTokenInfo: async () =>
            ok({
              scopes: [],
              audience: undefined,
              expiresAt: undefined,
              expiresInSeconds: 3600,
              elevated: { ...everyTier, refresh: 'interactive' as const },
              chatsvcagg: { ...everyTier, refresh: 'automatic' as const },
              ic3: { ...everyTier, refresh: 'automatic' as const },
            }),
        }),
    });
    const text = textOf(await client.callTool({ name: 'login', arguments: {} }));
    expect(text).toContain('available: [basic, elevated, chatsvcagg, ic3]');
  });
});
