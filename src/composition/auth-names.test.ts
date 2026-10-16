import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import type { AuthLadder } from '../infra/auth.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { fakeCommandPackage } from '@ask-marcel/office-core/test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { addAuthCommands } from './auth-cli.ts';
import { addAuthTools } from './auth-mcp.ts';
import { buildCliShell } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { buildMcpShell } from '@ask-marcel/office-core/composition/mcp-shell.ts';

// The auth package of a split (widget-auth; read is widget-cli, write is
// widget-write). Every text its login and status print names the auth bin, never
// a placeholder and never the single package's bin.

const AUTH_PACKAGE = fakeCommandPackage({ binName: 'widget-auth', bins: { auth: 'widget-auth', read: 'widget-cli', write: 'widget-write' } });

const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/;

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run().catch(() => undefined);
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const cliOf = (auth: AuthLadder): ReturnType<typeof buildCliShell> =>
  buildCliShell({
    commandPackage: AUTH_PACKAGE,
    graph: fakeGraphClient(),
    logger: createLoggerFake(),
    fs: createFileSystemFake(),
    version: '1.2.3',
    addLifecycleCommands: (program, output) => addAuthCommands(program, output, { auth, names: AUTH_PACKAGE }),
  });

const connect = async (auth: AuthLadder): Promise<Client> => {
  const server = buildMcpShell({
    commandPackage: AUTH_PACKAGE,
    graph: fakeGraphClient(),
    fs: createFileSystemFake(),
    version: '1.2.3',
    addTools: (mcp) => addAuthTools(mcp, { auth, names: AUTH_PACKAGE }),
  });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');

const refusing = (): AuthLadder =>
  fakeAuthManager({ getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'Refused. Run `{auth-bin} login --force`.' } }) });

describe('the bins the auth commands of a terminal name', () => {
  it('names the auth bin in the refresh advice of status, for a cache that misses three tokens', async () => {
    const out = await captureStdout(() => cliOf(fakeAuthManager()).parseAsync(['node', 'widget-auth', 'status']));
    expect(out).toContain('widget-auth login');
    expect(out).not.toMatch(PLACEHOLDER);
  });

  it('names the auth bin in the example of the status help', async () => {
    const help = await captureStdout(async () =>
      cliOf(fakeAuthManager())
        .commands.find((c) => c.name() === 'status')
        ?.outputHelp()
    );
    expect(help).toContain('Example:       widget-auth status');
    expect(help).not.toMatch(PLACEHOLDER);
  });
});

describe('the bins the auth tools of an MCP server name', () => {
  it('names the auth bin in the summary of a sign-in that left three tokens missing', async () => {
    const out = textOf(await (await connect(fakeAuthManager())).callTool({ name: 'login', arguments: {} }));
    expect(out).toContain('`widget-auth status`');
    expect(out).not.toMatch(PLACEHOLDER);
  });

  it('names the auth bin in the message of a sign-in that failed', async () => {
    const out = textOf(await (await connect(refusing())).callTool({ name: 'login', arguments: {} }));
    expect(out).toContain('error: Refused. Run `widget-auth login --force`.');
    expect(out).not.toMatch(PLACEHOLDER);
  });

  it('names the auth bin in the description of the login tool', async () => {
    const { tools } = await (await connect(fakeAuthManager())).listTools();
    const description = tools.find((t) => t.name === 'login')?.description ?? '';
    expect(description).toContain('`widget-auth login`');
    expect(description).not.toMatch(PLACEHOLDER);
  });
});
