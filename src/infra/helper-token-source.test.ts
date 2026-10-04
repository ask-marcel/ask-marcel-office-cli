import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import { createHelperTokenSource } from './helper-token-source.ts';
import type { HelperToken } from './token-helper-answer.ts';
import type { TokenHelperRequest } from './token-helper-run.ts';

const T0 = 1_800_000_000_000;
const MINUTE = 60_000;
const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
const inAnHour = Math.floor(T0 / 1000) + 3600;
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (name: string): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment({ name })}.sig`);

type Answer = Result<HelperToken, TokenError>;
const answered = (token: AccessToken, extra: Partial<HelperToken> = {}): Answer => ok({ token, expiresOn: inAnHour, ...extra });

// What a request asked the helper, as its command line would say it.
const asked = (request: TokenHelperRequest): string => [request.tier, request.tenant, request.rejected && `past ${request.rejected}`].filter(Boolean).join(' ');

type Setup = { readonly answer?: (request: TokenHelperRequest) => Answer | Promise<Answer>; readonly clock?: { now: number } };

// A helper source whose runs the test scripts, at a clock the test moves.
const sourceWith = (setup: Setup = {}): { source: ReturnType<typeof createHelperTokenSource>; runs: string[] } => {
  const clock = setup.clock ?? { now: T0 };
  const runs: string[] = [];
  const answer = setup.answer ?? ((request: TokenHelperRequest): Answer => answered(jwt(asked(request)), { region: 'emea' }));
  const ask = async (request: TokenHelperRequest): Promise<Answer> => {
    runs.push(asked(request));
    return answer(request);
  };
  return { source: createHelperTokenSource({ ask, now: () => clock.now }), runs };
};

describe('tokens from the token helper, kept in memory', () => {
  it('runs the helper once for nine concurrent calls on a cold source, and gives all nine the same token', async () => {
    const held = Promise.withResolvers<Answer>();
    const { source, runs } = sourceWith({ answer: () => held.promise });
    const pending = Array.from({ length: 9 }, () => source.graphToken('basic'));
    held.resolve(answered(jwt('shared')));
    expect(await Promise.all(pending)).toEqual(Array.from({ length: 9 }, () => ok(jwt('shared'))));
    expect(runs).toEqual(['basic']);
  });

  it('keeps one request in flight per tier and per tenant, never one for all', async () => {
    const other = tenantIdUnsafe('99999999-8888-7777-6666-555555555555');
    const { source, runs } = sourceWith();
    await Promise.all([source.graphToken('basic'), source.graphToken('elevated'), source.guestToken(TENANT), source.guestToken(other), source.guestToken(TENANT)]);
    expect(runs).toEqual(['basic', 'elevated', `guest ${TENANT}`, `guest ${other}`]);
  });

  it('reuses a token in memory for at most five minutes, then asks the helper again', async () => {
    const clock = { now: T0 };
    const { source, runs } = sourceWith({ clock });
    await source.graphToken('basic');
    clock.now = T0 + 5 * MINUTE - 1;
    await source.graphToken('basic');
    expect(runs).toHaveLength(1);
    clock.now = T0 + 5 * MINUTE;
    await source.graphToken('basic');
    expect(runs).toHaveLength(2);
  });

  it('never reuses a token past its expiry minus five minutes, even inside the five-minute window', async () => {
    const clock = { now: T0 };
    const { source, runs } = sourceWith({ clock, answer: () => answered(jwt('short-lived'), { expiresOn: Math.floor(T0 / 1000) + 7 * 60 }) });
    await source.graphToken('basic');
    clock.now = T0 + 2 * MINUTE - 1;
    await source.graphToken('basic');
    expect(runs).toHaveLength(1);
    clock.now = T0 + 2 * MINUTE;
    await source.graphToken('basic');
    expect(runs).toHaveLength(2);
  });

  it('uses a token whose expiresOn is 0 once and never reuses it, since 0 means already expired', async () => {
    const { source, runs } = sourceWith({ answer: () => answered(jwt('no-exp'), { expiresOn: 0 }) });
    expect(await source.graphToken('basic')).toEqual(ok(jwt('no-exp')));
    await source.graphToken('basic');
    expect(runs).toHaveLength(2);
  });

  it('asks for a guest token with the partner tenant', async () => {
    const { source, runs } = sourceWith();
    expect(await source.guestToken(TENANT)).toEqual(ok(jwt(`guest ${TENANT}`)));
    expect(runs).toEqual([`guest ${TENANT}`]);
  });

  it('takes the region from the same answer as the chat token, so a chat request runs the helper once', async () => {
    const { source, runs } = sourceWith({ answer: () => answered(jwt('ic3'), { region: 'amer' }) });
    expect(await source.substrateRegion('ic3')).toEqual(ok('amer' as TeamsRegion));
    expect(await source.substrateToken('ic3')).toEqual(ok(jwt('ic3')));
    expect(runs).toEqual(['ic3']);
  });

  it('refuses a chat answer that names no region, or one that is not a region name', async () => {
    for (const region of [undefined, 'emea/../admin']) {
      const { source } = sourceWith({ answer: () => answered(jwt('chat'), region === undefined ? {} : { region }) });
      const result = await source.substrateRegion('chatsvcagg');
      expect(result).toMatchObject({ ok: false, error: { type: 'auth_failed', code: 'secondary_token_unavailable' } });
      expect(JSON.stringify(result)).not.toContain('admin');
    }
  });

  it('replays a refused chat token past the token held in memory, by its fingerprint, and keeps the new one', async () => {
    const refused = jwt('refused');
    const { source, runs } = sourceWith({ answer: (request) => answered(request.rejected === undefined ? refused : jwt('fresh')) });
    expect(await source.substrateToken('chatsvcagg')).toEqual(ok(refused));
    expect(await source.substrateToken('chatsvcagg', { rejected: refused })).toEqual(ok(jwt('fresh')));
    expect(await source.substrateToken('chatsvcagg')).toEqual(ok(jwt('fresh')));
    expect(runs).toEqual(['chatsvcagg', `chatsvcagg past ${await tokenFingerprint(refused)}`]);
  });

  it('answers a replay with the token in memory when it is already a newer one than the refused token', async () => {
    const { source, runs } = sourceWith({ answer: () => answered(jwt('newer')) });
    await source.substrateToken('ic3');
    expect(await source.substrateToken('ic3', { rejected: jwt('older') })).toEqual(ok(jwt('newer')));
    expect(runs).toHaveLength(1);
  });

  it('runs one replay for concurrent callers refused with the same token', async () => {
    const { source, runs } = sourceWith();
    await Promise.all(Array.from({ length: 4 }, () => source.substrateToken('ic3', { rejected: jwt('refused') })));
    expect(runs).toEqual([`ic3 past ${await tokenFingerprint(jwt('refused'))}`]);
  });

  it('runs its own replay while a plain run is in flight, rather than share a run that may hand back the refused token', async () => {
    const plain = Promise.withResolvers<Answer>();
    const refused = jwt('refused');
    const { source, runs } = sourceWith({ answer: (request) => (request.rejected === undefined ? plain.promise : answered(jwt('fresh'))) });
    const first = source.substrateToken('ic3');
    expect(await source.substrateToken('ic3', { rejected: refused })).toEqual(ok(jwt('fresh')));
    plain.resolve(answered(refused));
    await first;
    expect(runs).toEqual(['ic3', `ic3 past ${await tokenFingerprint(refused)}`]);
  });

  it('does not let a plain run that ends after a replay put the refused token back in memory', async () => {
    const plain = Promise.withResolvers<Answer>();
    const refused = jwt('refused');
    const { source, runs } = sourceWith({ answer: (request) => (request.rejected === undefined ? plain.promise : answered(jwt('fresh'))) });
    const first = source.substrateToken('ic3');
    await source.substrateToken('ic3', { rejected: refused });
    plain.resolve(answered(refused));
    expect(await first).toEqual(ok(refused));
    expect(await source.substrateToken('ic3')).toEqual(ok(jwt('fresh')));
    expect(runs).toHaveLength(2);
  });

  it('keeps no failure in memory: the next call asks the helper again', async () => {
    let answers = 0;
    const failed: Result<never, TokenError> = err({ type: 'auth_failed', message: 'busy', code: 'sign_in_in_progress' });
    const { source, runs } = sourceWith({ answer: () => (++answers === 1 ? failed : answered(jwt('second'))) });
    expect(await source.graphToken('basic')).toEqual(failed);
    expect(await source.graphToken('basic')).toEqual(ok(jwt('second')));
    expect(runs).toHaveLength(2);
  });
});
