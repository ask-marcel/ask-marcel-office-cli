import { describe, expect, it } from 'bun:test';
import { tenantIdUnsafe } from '@ask-marcel/office-core/domain/tenant-id.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { LoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { FetchFn } from './auth.ts';
import { createAuthLadder } from './auth.ts';

// The token endpoint refuses the refresh token (a 4xx). Entra's answer names the
// cause in two codes, `error` and an AADSTS number, next to a trace id, a
// correlation id and a description. The redemption keeps the two codes only, in
// its message and in one warn event, so a dead refresh token is visible, and it
// quotes nothing else of the body.

const CACHE_PATH = '/virtual/token-cache.json';
const PARTNER = tenantIdUnsafe('8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f');
const LEAKED = 'eyJleaked-secret';
const TRACE = 'trace-5d1c7e2a';

const SESSION = JSON.stringify({ access_token: 'expired', expires_on: 0, refresh_token: 'the-refresh-token', chatsvcagg_region: 'emea' });

const SPA_LIFETIME_ENDED = JSON.stringify({
  error: 'invalid_grant',
  error_description: `AADSTS700084: The refresh token was issued to a single page app (SPA), and therefore has a fixed, limited lifetime of 1.00:00:00. Trace ID: ${TRACE} ${LEAKED}`,
  error_codes: [700084],
  trace_id: TRACE,
  correlation_id: 'corr-91b0',
});

const sessionOnDisk = (): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE_PATH, SESSION);
  return fs;
};

const refusing = (status: number, body: string): FetchFn => {
  return async () => new Response(body, { status, headers: { 'content-type': 'application/json' } });
};

const ladderOn = (fs: FileSystemFake, fetchFn: FetchFn, logger: LoggerFake = createLoggerFake()): ReturnType<typeof createAuthLadder> =>
  createAuthLadder({ cachePath: CACHE_PATH, browserProfileDir: '/virtual/profile', logger, fs, interactive: false, fetchFn });

const refusals = (logger: LoggerFake): LoggerFake['calls'] => logger.calls.filter((c) => c.event === 'auth.refresh.redeem_failed');

describe('a token endpoint that refuses the refresh token', () => {
  it('a guest redemption names the status, the error and the AADSTS number, and quotes nothing else of the body', async () => {
    const result = await ladderOn(sessionOnDisk(), refusing(400, SPA_LIFETIME_ENDED)).getGuestAccessToken(PARTNER);
    expect(result).toMatchObject({ ok: false, error: { message: expect.stringContaining('(refresh failed (400, invalid_grant, AADSTS700084))') } });
    expect(JSON.stringify(result)).not.toContain('leaked');
    expect(JSON.stringify(result)).not.toContain(TRACE);
  });

  it('a Teams chat history token redemption writes one warn event with the codes, and the user still gets the run-login message', async () => {
    const logger = createLoggerFake();
    const result = await ladderOn(sessionOnDisk(), refusing(400, SPA_LIFETIME_ENDED), logger).getIc3AccessToken();
    expect(result).toMatchObject({ ok: false, error: { code: 'secondary_token_unavailable' } });
    expect(refusals(logger)).toEqual([{ level: 'warn', event: 'auth.refresh.redeem_failed', meta: { tier: 'ic3', status: 400, error: 'invalid_grant', aadsts: 'AADSTS700084' } }]);
  });

  it('names the tier of each redemption in its event, and never the tenant of a guest call', async () => {
    const logger = createLoggerFake();
    const ladder = ladderOn(sessionOnDisk(), refusing(400, SPA_LIFETIME_ENDED), logger);
    await ladder.getAccessToken();
    await ladder.getChatsvcaggAccessToken();
    await ladder.getGuestAccessToken(PARTNER);
    expect(refusals(logger).map((c) => (c.meta as Record<string, unknown>)['tier'])).toEqual(['basic', 'chatsvcagg', 'guest']);
    expect(JSON.stringify(logger.calls)).not.toContain(PARTNER);
    expect(JSON.stringify(logger.calls)).not.toContain('leaked');
    expect(JSON.stringify(logger.calls)).not.toContain(TRACE);
  });

  it('keeps only the status when the refusal is not JSON, or carries no codes', async () => {
    for (const body of [`<html>${LEAKED}</html>`, '{}', JSON.stringify({ error: 42, error_description: 'no number here' })]) {
      const logger = createLoggerFake();
      const result = await ladderOn(sessionOnDisk(), refusing(401, body), logger).getGuestAccessToken(PARTNER);
      expect(result).toMatchObject({ ok: false, error: { message: expect.stringContaining('(refresh failed (401))') } });
      expect(refusals(logger)).toEqual([{ level: 'warn', event: 'auth.refresh.redeem_failed', meta: { tier: 'guest', status: 401 } }]);
    }
  });

  // An `error` that is not a plain lower-case code can be anything a proxy put
  // there, a token too: it is dropped, and the AADSTS number stays.
  it('drops an error field that is not a plain code', async () => {
    const body = JSON.stringify({ error: `invalid_grant ${LEAKED}`, error_description: 'AADSTS50173: the grant was revoked.' });
    const logger = createLoggerFake();
    const result = await ladderOn(sessionOnDisk(), refusing(400, body), logger).getGuestAccessToken(PARTNER);
    expect(result).toMatchObject({ ok: false, error: { message: expect.stringContaining('(refresh failed (400, AADSTS50173))') } });
    expect(JSON.stringify([result, logger.calls])).not.toContain('leaked');
  });
});
