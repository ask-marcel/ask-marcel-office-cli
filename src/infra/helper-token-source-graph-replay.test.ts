import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import { createHelperTokenSource } from './helper-token-source.ts';
import type { HelperToken } from './token-helper-answer.ts';
import type { TokenHelperRequest } from './token-helper-run.ts';

const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (name: string): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment({ name })}.sig`);

type Answer = Result<HelperToken, TokenError>;

// What a request asked the helper, as its command line would say it.
const asked = (request: TokenHelperRequest): string => [request.tier, request.tenant, request.rejected && `past ${request.rejected}`].filter(Boolean).join(' ');

// A helper that hands out `refused` until it is asked past a token, then `fresh`.
const sourceWith = (refused: AccessToken): { source: ReturnType<typeof createHelperTokenSource>; runs: string[] } => {
  const runs: string[] = [];
  const ask = async (request: TokenHelperRequest): Promise<Answer> => {
    runs.push(asked(request));
    return ok({ token: request.rejected === undefined ? refused : jwt('fresh'), expiresOn: inAnHour() });
  };
  return { source: createHelperTokenSource({ ask }), runs };
};

describe('a Graph token that Graph refused, replayed through the token helper', () => {
  it('asks the helper for a basic token past the refused one, by its fingerprint, and keeps the new one', async () => {
    const refused = jwt('refused');
    const { source, runs } = sourceWith(refused);
    expect(await source.graphToken('basic')).toEqual(ok(refused));
    expect(await source.graphToken('basic', { rejected: refused })).toEqual(ok(jwt('fresh')));
    expect(await source.graphToken('basic')).toEqual(ok(jwt('fresh')));
    expect(runs).toEqual(['basic', `basic past ${await tokenFingerprint(refused)}`]);
  });

  it('asks the helper for a guest token of the same partner tenant past the refused one', async () => {
    const refused = jwt('refused');
    const { source, runs } = sourceWith(refused);
    await source.guestToken(TENANT);
    expect(await source.guestToken(TENANT, { rejected: refused })).toEqual(ok(jwt('fresh')));
    expect(runs).toEqual([`guest ${TENANT}`, `guest ${TENANT} past ${await tokenFingerprint(refused)}`]);
  });
});
