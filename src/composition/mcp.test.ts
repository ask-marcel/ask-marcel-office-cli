import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { commands } from '../use-cases/commands/index.ts';
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

const readCommandCount = Object.values(commands).filter((c) => c.meta.mutates !== true).length;
const writeCommandNames = Object.entries(commands)
  .filter(([, c]) => c.meta.mutates === true)
  .map(([n]) => n);

describe('the MCP gateway an agent connects to', () => {
  it('offers exactly the five gateway tools, not one tool per command, so a session is not flooded with schema', async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['get-command-docs', 'list-commands', 'login', 'run-command', 'run-write-command']);
  });

  it('marks the read tools read-only so a client can auto-approve them, and does not mark the write tool read-only', async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    const annotationOf = (name: string): Record<string, unknown> => tools.find((t) => t.name === name)?.annotations ?? {};
    expect(annotationOf('list-commands')['readOnlyHint']).toBe(true);
    expect(annotationOf('get-command-docs')['readOnlyHint']).toBe(true);
    expect(annotationOf('run-command')['readOnlyHint']).toBe(true);
    expect(annotationOf('run-write-command')['readOnlyHint']).toBe(false);
    // The drafts are unsent and overwrite nothing.
    expect(annotationOf('run-write-command')['destructiveHint']).toBe(false);
  });
});

describe('discovering what the CLI can do over MCP', () => {
  it('lists every command when no category is given', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'list-commands', arguments: {} }));
    expect(text).toContain('list-mail-messages');
    expect(text).toContain('get-current-user');
  });

  it('narrows the manifest to one category so a fresh session can discover mail without paying for every command', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'list-commands', arguments: { category: 'mail' } }));
    expect(text).toContain('list-mail-messages');
    expect(text).not.toContain('list-calendars');
  });

  it('rejects an unknown category by naming the real ones instead of silently returning nothing', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'list-commands', arguments: { category: 'emails' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('mail');
  });

  it('returns a command`s full docs so an agent learns its options before invoking it', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'get-command-docs', arguments: { command: 'list-mail-messages' } }));
    expect(text).toContain('list-mail-messages');
    // The docs must carry the actual invocation surface, not just the name.
    expect(text).toContain('--top');
    expect(text).toContain('/me/messages');
  });

  it('returns docs for a command asked for by a deprecated name, since the manifest still advertises those names', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'get-command-docs', arguments: { command: 'download-onedrive-file-content' } }));
    expect(text).toContain('download-drive-item-content');
  });

  it('returns docs for a lifecycle command, which lives outside the registry', async () => {
    const client = await connect();
    const text = textOf(await client.callTool({ name: 'get-command-docs', arguments: { command: 'login' } }));
    expect(text).toContain('login');
  });

  it('rejects docs for an unknown command and points back at list-commands', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'get-command-docs', arguments: { command: 'list-mail-messagez' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('list-commands');
  });
});

describe('running a read command over MCP', () => {
  it('runs the command and returns its rendered result', async () => {
    const graph: GraphClient = fakeGraphClient({ get: async () => ok({ displayName: 'Robin Chen' }) });
    const client = await connect({ graph });
    const result = await client.callTool({ name: 'run-command', arguments: { command: 'get-current-user' } });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('Robin Chen');
  });

  it('passes params through without the `--` prefix an agent has no way to use', async () => {
    let capturedUrl = '';
    const graph: GraphClient = fakeGraphClient({
      get: async (url: string) => {
        capturedUrl = url;
        return ok({ value: [] });
      },
    });
    const client = await connect({ graph });
    await client.callTool({ name: 'run-command', arguments: { command: 'list-mail-messages', params: { top: '5' } } });
    expect(capturedUrl).toContain('$top=5');
  });

  it('surfaces a Graph failure as a tool error carrying the curated hint, so the agent can self-correct', async () => {
    const graph: GraphClient = fakeGraphClient({
      get: async () => ({ ok: false, error: { type: 'api_error', status: 404, message: 'itemNotFound', code: 'itemNotFound' } }) as never,
    });
    const client = await connect({ graph });
    const result = await client.callTool({ name: 'run-command', arguments: { command: 'get-current-user' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('error:');
  });

  it('rejects an unknown command instead of failing somewhere deeper', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'run-command', arguments: { command: 'no-such-command' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('list-commands');
  });

  it('runs a command asked for by its deprecated name', async () => {
    const graph: GraphClient = fakeGraphClient({ getBinary: async () => ok({ contentType: 'text/plain', size: 2, base64: 'aGk=' }) });
    const client = await connect({ graph });
    const result = await client.callTool({
      name: 'run-command',
      arguments: { command: 'download-onedrive-file-content', params: { driveId: 'd1', itemId: 'i1' } },
    });
    expect(isError(result)).toBe(false);
  });
});

describe('the read/write boundary the readOnlyHint promises', () => {
  it.each(writeCommandNames)('refuses to run the write command `%s` through the read-only run-command tool', async (writeName) => {
    // The whole point of splitting the tools: a client may auto-approve
    // run-command on the strength of readOnlyHint: true. A write slipping
    // through here would mutate under that promise.
    const client = await connect();
    const result = await client.callTool({ name: 'run-command', arguments: { command: writeName } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('run-write-command');
  });

  it('refuses to run a read command through run-write-command, keeping each tool`s advertised set honest', async () => {
    const client = await connect();
    const result = await client.callTool({ name: 'run-write-command', arguments: { command: 'get-current-user' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('run-command');
  });

  it('rejects the write command BEFORE executing it, so a refusal can never leave a draft behind', async () => {
    let graphWasCalled = false;
    const graph: GraphClient = fakeGraphClient({
      post: async () => {
        graphWasCalled = true;
        return ok({});
      },
    });
    const client = await connect({ graph });
    await client.callTool({
      name: 'run-command',
      arguments: { command: 'create-mail-draft', params: { subject: 'x', bodyContent: 'y', toRecipients: 'robin.chen@contoso.com' } },
    });
    expect(graphWasCalled).toBe(false);
  });

  it('runs a write command through run-write-command', async () => {
    const graph: GraphClient = fakeGraphClient({ post: async () => ok({ id: 'draft-1', subject: 'Hello' }) });
    const client = await connect({ graph });
    const result = await client.callTool({
      name: 'run-write-command',
      arguments: { command: 'create-mail-draft', params: { subject: 'Hello', bodyContent: 'Hi', toRecipients: 'robin.chen@contoso.com' } },
    });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('draft-1');
  });

  it('advertises the write set derived from the registry, so a newly-flagged write command is covered without touching this file', async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    const description = tools.find((t) => t.name === 'run-write-command')?.description ?? '';
    for (const name of writeCommandNames) expect(description).toContain(name);
    expect(tools.find((t) => t.name === 'run-command')?.description).toContain(String(readCommandCount));
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
      graph: fakeGraphClient({ getCachedTokenInfo: async () => ({ ok: false, error: { type: 'auth_failed', message: 'cache unreadable' } }) as never }),
    });
    const result = await client.callTool({ name: 'login', arguments: {} });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('cache unreadable');
  });
});
