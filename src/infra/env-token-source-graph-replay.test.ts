import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { createEnvTokenSource } from './env-token-source.ts';

const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment(claims)}.sig`);
const GRAPH = 'https://graph.microsoft.com';
const HELPER_TOKEN = jwt({ exp: inAnHour(), aud: GRAPH, from: 'helper' });

type HelperFake = TokenSource & { readonly asked: ReadonlyArray<string> };

// The token helper behind the environment: it records each ask, and whether the
// ask was past a refused token.
const helperFake = (): HelperFake => {
  const asked: string[] = [];
  const past = (what: string, rejected: AccessToken | undefined): string => (rejected === undefined ? what : `${what} past a refused token`);
  return {
    asked,
    graphToken: async (tier, options) => {
      asked.push(past(tier, options?.rejected));
      return ok(HELPER_TOKEN);
    },
    guestToken: async (tenant, options) => {
      asked.push(past(`guest ${tenant}`, options?.rejected));
      return ok(HELPER_TOKEN);
    },
    substrateToken: async () => ok(HELPER_TOKEN),
    substrateRegion: async () => ok('amer' as TeamsRegion),
  };
};

describe('a Graph token from the environment that Graph refused', () => {
  it('surfaces a refused ASKMARCEL_TOKEN_BASIC as env_token_invalid, never quoting it, and never replays it with a token from the helper', async () => {
    const token = jwt({ exp: inAnHour(), aud: GRAPH });
    const helper = helperFake();
    const refused = await createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: token }, helper).graphToken('basic', { rejected: token });
    expect(refused.ok).toBe(false);
    if (refused.ok || refused.error.type !== 'auth_failed') return;
    expect(refused.error.code).toBe('env_token_invalid');
    expect(refused.error.message).toContain('Microsoft Graph refused the token in ASKMARCEL_TOKEN_BASIC (HTTP 401)');
    expect(refused.error.message).not.toContain(token);
    expect(helper.asked).toEqual([]);
  });

  it('lets the helper replay a refused basic token when ASKMARCEL_TOKEN_BASIC is not set', async () => {
    const helper = helperFake();
    expect(await createEnvTokenSource({}, helper).graphToken('basic', { rejected: jwt({ exp: inAnHour(), aud: GRAPH }) })).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual(['basic past a refused token']);
  });

  it('lets the helper replay a refused guest token, since guest tokens have no environment form', async () => {
    const helper = helperFake();
    const tenant = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: jwt({ exp: inAnHour(), aud: GRAPH }) }, helper);
    expect(await source.guestToken(tenant, { rejected: jwt({ exp: inAnHour(), aud: GRAPH }) })).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual([`guest ${tenant} past a refused token`]);
  });

  it('names the service that refused a token from each variable', async () => {
    const helper = helperFake();
    const token = jwt({ exp: inAnHour(), aud: GRAPH });
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_ELEVATED: token, ASKMARCEL_TOKEN_CHATSVCAGG: token, ASKMARCEL_TOKEN_IC3: token }, helper);
    const messages = await Promise.all([
      source.graphToken('elevated', { rejected: token }),
      source.substrateToken('chatsvcagg', { rejected: token }),
      source.substrateToken('ic3', { rejected: token }),
    ]);
    expect(messages.map((r) => (r.ok || r.error.type !== 'auth_failed' ? '' : r.error.message.slice(0, r.error.message.indexOf(' (HTTP 401)'))))).toEqual([
      'Microsoft Graph refused the token in ASKMARCEL_TOKEN_ELEVATED',
      'The Teams chat service refused the token in ASKMARCEL_TOKEN_CHATSVCAGG',
      'The Teams chat service refused the token in ASKMARCEL_TOKEN_IC3',
    ]);
  });
});
