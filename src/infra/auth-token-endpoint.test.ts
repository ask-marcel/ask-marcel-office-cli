import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '../test-helpers/jwt.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { FetchFn } from './auth.ts';
import { createAuthLadder } from './auth.ts';

// The token endpoint answers 200, but its body is not a token answer: a captive
// portal's page, a proxy's error, a half-written JSON. The redemption fails as a
// Result that names the status and quotes nothing of the body, which can carry a
// token, and the refresh token on disk stays as it was.

const CACHE_PATH = '/virtual/token-cache.json';
const PARTNER = tenantIdUnsafe('8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f');
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(unsignedJwt(claims));
const LEAKED = 'eyJleaked-secret';

const SESSION = JSON.stringify({ access_token: 'expired', expires_on: 0, refresh_token: 'the-refresh-token', chatsvcagg_region: 'emea' });

const sessionOnDisk = (): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, SESSION);
  return fs;
};

const answering200 = (body: string): FetchFn => {
  return async () => new Response(body, { status: 200, headers: { 'content-type': 'application/json' } });
};

const ladderOn = (fs: FileSystemFake, fetchFn: FetchFn): ReturnType<typeof createAuthLadder> =>
  createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger: createLoggerFake(), fs, interactive: false, fetchFn });

const NOT_JSON = `<html><body>Sign in again. access_token=${LEAKED}</body></html>`;

describe('a token endpoint that answers 200 with something other than a token', () => {
  it('a Graph refresh fails as not signed in, quotes nothing of the body, and keeps the refresh token on disk', async () => {
    const fs = sessionOnDisk();
    const result = await ladderOn(fs, answering200(NOT_JSON)).getAccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'not_authenticated' } });
    expect(JSON.stringify(result)).not.toContain('leaked');
    expect(fs.snapshot(CACHE_PATH)).toBe(SESSION);
  });

  // The two chat tiers fall back to their own "run login" message, as on any
  // failed redemption.
  it('a Teams chat or chat history token redemption fails as unavailable, quotes nothing of the body, and keeps the refresh token', async () => {
    for (const tier of ['chatsvcagg', 'ic3'] as const) {
      const fs = sessionOnDisk();
      const ladder = ladderOn(fs, answering200(NOT_JSON));
      const result = tier === 'chatsvcagg' ? await ladder.getChatsvcaggAccessToken() : await ladder.getIc3AccessToken();
      expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
      expect(JSON.stringify(result)).not.toContain('leaked');
      expect(fs.snapshot(CACHE_PATH)).toBe(SESSION);
    }
  });

  it('a guest token redemption names the tenant and the status, and quotes nothing of the body', async () => {
    const fs = sessionOnDisk();
    const result = await ladderOn(fs, answering200(NOT_JSON)).getGuestAccessToken(PARTNER);
    const message = `tenant ${PARTNER} refused a guest token (refresh failed (200): its answer could not be read as JSON)`;
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable', message: expect.stringContaining(message) } });
    expect(JSON.stringify(result)).not.toContain('leaked');
    expect(fs.snapshot(CACHE_PATH)).toBe(SESSION);
  });

  // The redemption's deadline also covers the body: headers arrive, then the
  // body stalls and the deadline fires while it is read. That is a slow
  // endpoint, not a malformed answer.
  it('an answer whose body stalls past the deadline is reported as timed out, quotes nothing of the body, and keeps the refresh token', async () => {
    const stalled: FetchFn = async () => {
      const body = new ReadableStream({
        start: (controller) => {
          controller.enqueue(new TextEncoder().encode(`{"access_token":"${LEAKED}`));
          controller.error(new DOMException(`timed out reading ${LEAKED}`, 'TimeoutError'));
        },
      });
      return new Response(body, { status: 200, headers: { 'content-type': 'application/json' } });
    };
    const fs = sessionOnDisk();
    const result = await ladderOn(fs, stalled).getGuestAccessToken(PARTNER);
    const message = `tenant ${PARTNER} refused a guest token (refresh failed (200): its answer did not arrive within 60s)`;
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable', message: expect.stringContaining(message) } });
    expect(JSON.stringify(result)).not.toContain('leaked');
    expect(fs.snapshot(CACHE_PATH)).toBe(SESSION);
  });

  it('refuses JSON that is not shaped as a token answer, for every tier, and saves nothing', async () => {
    const minted = jwt({ exp: inAnHour(), aud: 'https://chatsvcagg.teams.microsoft.com', note: LEAKED });
    const shapes = [
      'null',
      '[]',
      JSON.stringify(LEAKED),
      JSON.stringify({ access_token: 42 }),
      JSON.stringify({ access_token: minted, expires_in: 'soon', refresh_token: LEAKED }),
      JSON.stringify({ access_token: minted, expires_in: 3600, refresh_token: 7 }),
      JSON.stringify({ access_token: minted, expires_in: 3600, refresh_token: null }),
    ];
    for (const body of shapes) {
      const fs = sessionOnDisk();
      const ladder = ladderOn(fs, answering200(body));
      const graph = await ladder.getAccessToken();
      const chat = await ladder.getChatsvcaggAccessToken();
      const guest = await ladder.getGuestAccessToken(PARTNER);
      expect(graph).toMatchObject({ ok: false, error: { code: 'not_authenticated' } });
      expect(chat).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
      expect(guest).toMatchObject({ ok: false, error: { message: expect.stringContaining('(refresh failed (200): its answer is not a token response)') } });
      expect(JSON.stringify([graph, chat, guest])).not.toContain('leaked');
      expect(fs.snapshot(CACHE_PATH)).toBe(SESSION);
    }
  });
});
