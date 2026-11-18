import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { AuthLadder } from '../infra/auth.ts';
import type { TokenInfo } from '../use-cases/ports/token-report.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { buildAuthCli } from './auth-program.ts';
import { buildAuthMcpServer } from './auth-server.ts';

// A sign-in that leaves no refresh token in the token cache is not a good
// session: its tokens stop at their expiry. `login` then fails with the code
// session_cannot_renew and the advice `login --force`, on the CLI and on MCP.

const MESSAGE =
  'The sign-in is complete, but the token cache holds no refresh token. Thus the session cannot renew itself: the basic, chatsvcagg and ic3 tokens stop at their expiry. Run `ask-marcel-office-auth login --force` to capture a new refresh token. It clears the browser session, so you can be asked for your password.';

const every = { available: true, expiresInSeconds: 3600, scopes: [] };
const info = (refreshTokenCached: boolean): TokenInfo => ({
  scopes: [],
  audience: undefined,
  expiresAt: undefined,
  expiresInSeconds: 3600,
  refreshTokenCached,
  elevated: { ...every, refresh: 'interactive' },
  chatsvcagg: { ...every, refresh: 'automatic' },
  ic3: { ...every, refresh: 'automatic' },
});
const signedIn = (refreshTokenCached: boolean): AuthLadder => fakeAuthManager({ getTokenInfo: async () => ok(info(refreshTokenCached)) });

const loginOnCli = async (auth: AuthLadder): Promise<{ readonly ok: boolean; readonly errorCode?: string; readonly error?: string; readonly data?: unknown }> => {
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await buildAuthCli({ auth, logger: createLoggerFake(), fs: createFileSystemFake() }).parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']);
  } finally {
    process.stdout.write = original;
  }
  return JSON.parse(out.trim()) as { ok: boolean; errorCode?: string; error?: string };
};

const loginOnMcp = async (auth: AuthLadder): Promise<{ readonly isError: boolean; readonly text: string }> => {
  const server = buildAuthMcpServer({ auth, version: '9.9.9' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  const result = (await client.callTool({ name: 'login', arguments: {} })) as { isError?: boolean; content: ReadonlyArray<{ text?: string }> };
  return { isError: result.isError === true, text: result.content.map((c) => c.text ?? '').join('') };
};

describe('login with no refresh token in the token cache after the sign-in', () => {
  it('fails on the CLI with the code session_cannot_renew and the advice login --force', async () => {
    const answer = await loginOnCli(signedIn(false));
    expect(answer).toMatchObject({ ok: false, errorCode: 'session_cannot_renew', error: MESSAGE });
  });

  it('fails on MCP with the same message and code', async () => {
    const answer = await loginOnMcp(signedIn(false));
    expect(answer.isError).toBe(true);
    expect(answer.text).toBe(`error: ${MESSAGE}\ncode: session_cannot_renew\n`);
  });
});

describe('login with the refresh token in the token cache', () => {
  it('reports the session as authenticated on the CLI, as before', async () => {
    expect(await loginOnCli(signedIn(true))).toMatchObject({ ok: true, data: { status: 'authenticated' } });
  });

  it('reports the session as authenticated on MCP, as before', async () => {
    const answer = await loginOnMcp(signedIn(true));
    expect(answer.isError).toBe(false);
    expect(answer.text).toContain('status: authenticated');
  });
});
