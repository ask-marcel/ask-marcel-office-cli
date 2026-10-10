import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import type { TokenIssuer } from '../use-cases/ports/token-issuer.ts';
import { createAuthManagerTokenSource } from './auth-token-source.ts';
import { createGraphClient } from './graph-client.ts';
import type { FetchFn } from './graph-request.ts';

const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (name: string): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment({ name })}.sig`);
const REFUSED = jwt('refused');

// The in-process ladder: its plain getters hand out the refused token, and its
// issuer records each ask past a refused token and answers a fresh one.
const ladder = (issued: string[]): ReturnType<typeof fakeAuthManager> & Pick<TokenIssuer, 'issueToken'> => ({
  ...fakeAuthManager({
    getAccessToken: async () => ok(REFUSED),
    getElevatedAccessToken: async () => ok(REFUSED),
    getGuestAccessToken: async () => ok(REFUSED),
  }),
  issueToken: async (request, rejected) => {
    issued.push(`${request.tier} past ${rejected}`);
    return ok(jwt('fresh'));
  },
});

describe('a Graph token that Graph refused, replayed through the in-process ladder', () => {
  it('asks the ladder for a basic token past the refused one by its fingerprint, which never opens a browser', async () => {
    const issued: string[] = [];
    const source = createAuthManagerTokenSource(ladder(issued));
    expect(await source.graphToken('basic', { rejected: REFUSED })).toEqual(ok(jwt('fresh')));
    expect(issued).toEqual([`basic past ${await tokenFingerprint(REFUSED)}`]);
  });

  it('asks the ladder for a guest token of the same partner tenant past the refused one', async () => {
    const issued: string[] = [];
    const source = createAuthManagerTokenSource(ladder(issued));
    expect(await source.guestToken(TENANT, { rejected: REFUSED })).toEqual(ok(jwt('fresh')));
    expect(issued).toEqual([`guest past ${await tokenFingerprint(REFUSED)}`]);
  });

  it('takes a plain basic or guest token from the getters when nothing was refused', async () => {
    const issued: string[] = [];
    const source = createAuthManagerTokenSource(ladder(issued));
    expect(await source.graphToken('basic')).toEqual(ok(REFUSED));
    expect(await source.guestToken(TENANT)).toEqual(ok(REFUSED));
    expect(issued).toEqual([]);
  });

  it('asks a manager with no issuer (a library caller bringing its own) its own getter again, which may answer the same token', async () => {
    const source = createAuthManagerTokenSource(fakeAuthManager({ getAccessToken: async () => ok(REFUSED), getGuestAccessToken: async () => ok(REFUSED) }));
    expect(await source.graphToken('basic', { rejected: REFUSED })).toEqual(ok(REFUSED));
    expect(await source.guestToken(TENANT, { rejected: REFUSED })).toEqual(ok(REFUSED));
  });

  it('sends a refused request once more when a manager with no issuer answers a different token from its own getter, and never more', async () => {
    const handed = [jwt('first'), jwt('second')];
    const auth = fakeAuthManager({ getAccessToken: async () => ok(handed.shift() ?? jwt('third')) });
    const seen: string[] = [];
    const fetchFn: FetchFn = async (_url, init) => {
      const bearer = new Headers(init?.headers).get('authorization') ?? '';
      seen.push(bearer);
      return Response.json({ error: { code: 'InvalidAuthenticationToken', message: 'expired' } }, { status: 401 });
    };
    const result = await createGraphClient(auth, fetchFn).get('/me');
    expect(result.ok).toBe(false);
    expect(seen).toEqual([`Bearer ${jwt('first')}`, `Bearer ${jwt('second')}`]);
  });

  it('never replays the elevated tier: it has no refresh token, so only a browser could renew it', async () => {
    const issued: string[] = [];
    const source = createAuthManagerTokenSource(ladder(issued));
    expect(await source.graphToken('elevated', { rejected: REFUSED })).toEqual(ok(REFUSED));
    expect(issued).toEqual([]);
  });
});
