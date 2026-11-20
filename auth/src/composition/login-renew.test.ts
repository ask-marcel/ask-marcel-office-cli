import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { AuthLadder } from '../infra/auth.ts';
import type { TokenInfo } from '../use-cases/ports/token-report.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { buildAuthCli } from './auth-program.ts';
import { buildAuthMcpServer } from './auth-server.ts';

// `login --renew` on the CLI and `renew: true` on the MCP `login` tool start a
// new sign-in session of about 24 hours, with the browser cookies kept. The
// answer of a successful login gives the end of the session on both.

type SignInOptions = { readonly force?: boolean; readonly renew?: boolean } | undefined;

const SESSION = { endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds: 86_390 };
const every = { available: true, expiresInSeconds: 3600, scopes: [] };
const info: TokenInfo = {
  scopes: [],
  audience: undefined,
  expiresAt: undefined,
  expiresInSeconds: 3600,
  refreshTokenCached: true,
  session: SESSION,
  elevated: { ...every, refresh: 'interactive' },
  chatsvcagg: { ...every, refresh: 'automatic' },
  ic3: { ...every, refresh: 'automatic' },
};

const recording = (): AuthLadder & { readonly signIns: SignInOptions[] } => {
  const signIns: SignInOptions[] = [];
  const auth = fakeAuthManager({
    getAccessToken: async (options) => {
      signIns.push(options);
      return ok(accessTokenUnsafe('tok'));
    },
    getTokenInfo: async () => ok(info),
  });
  return Object.assign(auth, { signIns });
};

const onCli = async (auth: AuthLadder, argv: ReadonlyArray<string>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await buildAuthCli({ auth, logger: createLoggerFake(), fs: createFileSystemFake() }).parseAsync(['node', 'ask-marcel-office', ...argv]);
  } finally {
    process.stdout.write = original;
  }
  return out;
};

const mcpClient = async (auth: AuthLadder): Promise<Client> => {
  const server = buildAuthMcpServer({ auth, version: '9.9.9' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');

describe('login --renew on the CLI', () => {
  it('asks the sign-in for the capture with the cookies kept', async () => {
    const auth = recording();
    await onCli(auth, ['--output', 'json', 'login', '--renew']);
    expect(auth.signIns).toEqual([{ force: false, renew: true }]);
  });

  it('answers with the end of the new session', async () => {
    const answer = JSON.parse(await onCli(recording(), ['--output', 'json', 'login', '--renew'])) as { data: { session: unknown } };
    expect(answer.data.session).toEqual(SESSION);
  });

  it('asks a plain login for no renew', async () => {
    const auth = recording();
    await onCli(auth, ['--output', 'json', 'login']);
    expect(auth.signIns[0]).toEqual({ force: false, renew: false });
  });
});

describe('the login tool of the MCP server', () => {
  it('takes renew and passes it to the sign-in', async () => {
    const auth = recording();
    const client = await mcpClient(auth);
    await client.callTool({ name: 'login', arguments: { renew: true } });
    expect(auth.signIns).toEqual([{ force: false, renew: true }]);
  });

  it('declares renew as an optional boolean', async () => {
    const { tools } = await (await mcpClient(recording())).listTools();
    const schema = tools.find((t) => t.name === 'login')?.inputSchema as { properties: Record<string, { type: string }>; required?: string[] };
    expect(schema.properties['renew']?.type).toBe('boolean');
    expect(schema.required ?? []).not.toContain('renew');
  });

  it('answers with the end of the session', async () => {
    const text = textOf(await (await mcpClient(recording())).callTool({ name: 'login', arguments: {} }));
    expect(text).toContain('2026-11-21T20:09:00.000Z');
  });
});
