import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import type { AuthLadder } from '../test-helpers/auth-manager-fake.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { buildAuthMcpServer } from './auth-server.ts';

// The MCP server of the auth package: its own server, not the gateway of the
// Graph commands, with the `login` and `status` tools only (LESSONS 2026-10-15).

const connect = async (auth: AuthLadder = fakeAuthManager()): Promise<Client> => {
  const server = buildAuthMcpServer({ auth, version: '1.0.0' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => {
  const content = (result as { content: ReadonlyArray<{ type: string; text?: string }> }).content;
  return content.map((c) => c.text ?? '').join('');
};

const inAnHour = Math.floor(Date.now() / 1000) + 3600;
const TOKENS = ['basic', 'elevated', 'chatsvcagg', 'ic3', 'guest'].map((tier) => unsignedJwt({ aud: 'https://graph.microsoft.com', exp: inAnHour, tier }));
const available = { available: true, expiresInSeconds: 3600, scopes: ['User.Read'], refresh: 'automatic' as const };

// A ladder that holds a real-shaped token for every tier.
const signedInAuth = (): AuthLadder =>
  fakeAuthManager({
    getAccessToken: async () => ok(accessTokenUnsafe(TOKENS[0] ?? '')),
    getElevatedAccessToken: async () => ok(accessTokenUnsafe(TOKENS[1] ?? '')),
    getChatsvcaggAccessToken: async () => ok(accessTokenUnsafe(TOKENS[2] ?? '')),
    getIc3AccessToken: async () => ok(accessTokenUnsafe(TOKENS[3] ?? '')),
    getGuestAccessToken: async () => ok(accessTokenUnsafe(TOKENS[4] ?? '')),
    getTokenInfo: async () =>
      ok({
        scopes: ['User.Read'],
        audience: 'https://graph.microsoft.com',
        expiresAt: undefined,
        expiresInSeconds: 3600,
        elevated: available,
        chatsvcagg: available,
        ic3: available,
      }),
  });

describe('the MCP server of the auth package', () => {
  it('offers exactly the login and status tools', async () => {
    const { tools } = await (await connect()).listTools();
    expect(tools.map((tool) => tool.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['login', 'status']);
  });

  it('names itself after the auth bin, with the version it is given', async () => {
    expect((await connect()).getServerVersion()).toEqual({ name: 'ask-marcel-office-auth', version: '1.0.0' });
  });

  it('names the auth bin in the report of the status tool', async () => {
    expect(textOf(await (await connect()).callTool({ name: 'status', arguments: {} }))).toContain('ask-marcel-office-auth login');
  });

  it('never puts a token in the result of a tool', async () => {
    const client = await connect(signedInAuth());
    const texts = [
      textOf(await client.callTool({ name: 'login', arguments: {} })),
      textOf(await client.callTool({ name: 'login', arguments: { force: true } })),
      textOf(await client.callTool({ name: 'status', arguments: {} })),
    ];
    expect(texts.every((text) => text.length > 0)).toBe(true);
    expect(texts.filter((text) => TOKENS.some((token) => text.includes(token) || text.includes(token.split('.')[1] ?? token)))).toEqual([]);
  });
});
