import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { tenantIdUnsafe } from '@ask-marcel/office-core/domain/tenant-id.ts';
import { tokenFingerprint } from '@ask-marcel/office-core/domain/token-fingerprint.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { BrowserRungs, FetchFn, TokenCacheAccess } from './auth.ts';
import { createAuthLadder } from './auth.ts';

// Entra refuses the refresh token with AADSTS700084 when the sign-in session
// ended (about 24 hours after the browser sign-in). The error then names the
// end of the session and a plain `login`, and carries the cause
// `session_ended`, so the token helper (whose logger is silent) can tell it.
// The codes stay: `not_authenticated` for the basic token,
// `secondary_token_unavailable` for the others.

const CACHE_PATH = '/virtual/token-cache.json';
const PARTNER = tenantIdUnsafe('8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f');
const LEAKED = 'eyJleaked-secret';
const now = (): number => Math.floor(Date.now() / 1000);

const SESSION_ENDED = JSON.stringify({
  error: 'invalid_grant',
  error_description: `AADSTS700084: The refresh token was issued to a single page app (SPA), and therefore has a fixed, limited lifetime of 1.00:00:00. ${LEAKED}`,
});
const OTHER_REFUSAL = JSON.stringify({ error: 'invalid_grant', error_description: 'AADSTS50173: The provided grant has expired.' });

const graphToken = (): string => unsignedJwt({ exp: now() + 3600, aud: 'https://graph.microsoft.com' });

const sessionOnDisk = (extra: Record<string, unknown> = {}): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired', expires_on: 0, refresh_token: 'the-refresh-token', chatsvcagg_region: 'emea', ...extra }));
  return fs;
};

const refusing = (body: string): FetchFn => {
  return async () => new Response(body, { status: 400, headers: { 'content-type': 'application/json' } });
};

const ladderOn = (fs: FileSystemFake, fetchFn: FetchFn, browserRungs?: (cache: TokenCacheAccess) => BrowserRungs): ReturnType<typeof createAuthLadder> =>
  createAuthLadder({
    cachePath: CACHE_PATH,
    browserProfileDir: '/virtual/profile',
    logger: createLoggerFake(),
    fs,
    interactive: false,
    fetchFn,
    ...(browserRungs ? { browserRungs } : {}),
  });

const SESSION_ENDED_WORDS = ['The sign-in session ended', 'about 24 hours', 'Run `{auth-bin} login`'];

const expectSessionEnded = (result: unknown, code: string): void => {
  expect(result).toMatchObject({ ok: false, error: { type: 'auth_failed', code, cause: 'session_ended' } });
  const message = (result as { error: { message: string } }).error.message;
  for (const words of SESSION_ENDED_WORDS) expect(message).toContain(words);
  expect(message).not.toContain('leaked');
};

describe('a redemption that Entra refuses because the sign-in session ended', () => {
  it('gives the basic token a not_authenticated error that names the end of the session and a plain login', async () => {
    expectSessionEnded(await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).getAccessToken(), 'not_authenticated');
  });

  it('gives the Teams chat token a secondary_token_unavailable error that names the end of the session', async () => {
    const result = await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).getChatsvcaggAccessToken();
    expectSessionEnded(result, 'secondary_token_unavailable');
    expect(JSON.stringify(result)).toContain('chatsvcagg');
  });

  it('gives the Teams chat history token the same error', async () => {
    expectSessionEnded(await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).getIc3AccessToken(), 'secondary_token_unavailable');
  });

  it('gives a guest token the end of the session, not a refusal of the partner tenant', async () => {
    const result = await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).getGuestAccessToken(PARTNER);
    expectSessionEnded(result, 'secondary_token_unavailable');
    expect(JSON.stringify(result)).not.toContain('you may not be a guest');
  });

  it('gives the same error to a replay of a refused basic token', async () => {
    const refused = await tokenFingerprint('a-refused-token');
    expectSessionEnded(await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).issueToken({ tier: 'basic' }, refused), 'not_authenticated');
  });

  it('gives the same error to a replay of a refused Teams chat token', async () => {
    const refused = await tokenFingerprint('a-refused-token');
    expectSessionEnded(await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED)).issueToken({ tier: 'ic3' }, refused), 'secondary_token_unavailable');
  });

  it('still opens the browser sign-in when this session may open one', async () => {
    let signIns = 0;
    const token = accessTokenUnsafe(graphToken());
    const rungs = (_cache: TokenCacheAccess): BrowserRungs => ({
      signIn: async () => {
        signIns += 1;
        return ok(token);
      },
      close: async () => {},
      lastElevatedOutcome: () => null,
      lastChatsvcaggOutcome: () => null,
    });
    const result = await ladderOn(sessionOnDisk(), refusing(SESSION_ENDED), rungs).getAccessToken();
    expect(result).toEqual(ok(token));
    expect(signIns).toBe(1);
  });
});

describe('a redemption that Entra refuses for another cause', () => {
  it('keeps the run-login error, with no cause', async () => {
    const result = await ladderOn(sessionOnDisk(), refusing(OTHER_REFUSAL)).getIc3AccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
    expect(JSON.stringify(result)).not.toContain('session_ended');
    expect(JSON.stringify(result)).not.toContain('about 24 hours');
  });
});

describe('the predicted end of the sign-in session', () => {
  it('does not stop a redemption: a cache past its end still sends the refresh token to Entra', async () => {
    const calls: string[] = [];
    const answering: FetchFn = async (url) => {
      calls.push(url);
      return new Response(JSON.stringify({ access_token: graphToken(), expires_in: 3600, refresh_token: 'rotated' }));
    };
    const result = await ladderOn(sessionOnDisk({ signed_in_at: now() - 25 * 3600 }), answering).getAccessToken();
    expect(result.ok).toBe(true);
    expect(calls).toHaveLength(1);
  });
});
