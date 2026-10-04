import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import { createEnvTokenSource } from './env-token-source.ts';

const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment(claims)}.sig`);
const GRAPH = 'https://graph.microsoft.com';
const CHATSVCAGG = 'https://chatsvcagg.teams.microsoft.com';
const IC3 = 'https://ic3.teams.office.com';

const HELPER_TOKEN = jwt({ exp: inAnHour(), aud: GRAPH, from: 'helper' });

type HelperFake = TokenSource & { readonly asked: ReadonlyArray<string> };

// The token helper behind the environment: it answers every tier and records
// what it was asked, so a test can prove it was never reached.
const helperFake = (): HelperFake => {
  const asked: string[] = [];
  return {
    asked,
    graphToken: async (tier) => {
      asked.push(tier);
      return ok(HELPER_TOKEN);
    },
    guestToken: async (tenant) => {
      asked.push(`guest ${tenant}`);
      return ok(HELPER_TOKEN);
    },
    substrateToken: async (tier, options) => {
      asked.push(options?.rejected === undefined ? tier : `${tier} past a refused token`);
      return ok(HELPER_TOKEN);
    },
    substrateRegion: async (tier) => {
      asked.push(`${tier} region`);
      return ok('amer' as TeamsRegion);
    },
  };
};

// The env_token_invalid failure a refused variable gives, or a test failure.
const invalid = (result: Result<unknown, TokenError>): string => {
  if (result.ok || result.error.type !== 'auth_failed' || result.error.code !== 'env_token_invalid') throw new Error('expected env_token_invalid');
  return result.error.message;
};

describe('tokens from the environment', () => {
  it('signs Graph calls with the token in ASKMARCEL_TOKEN_BASIC and never asks the token helper', async () => {
    const token = jwt({ exp: inAnHour(), aud: GRAPH });
    const helper = helperFake();
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: token }, helper);
    expect(await source.graphToken('basic')).toEqual(ok(token));
    expect(helper.asked).toEqual([]);
  });

  it('takes the elevated token from ASKMARCEL_TOKEN_ELEVATED, accepting Graph named by its application id', async () => {
    const token = jwt({ exp: inAnHour(), aud: '00000003-0000-0000-c000-000000000000' });
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_ELEVATED: token }, helperFake());
    expect(await source.graphToken('elevated')).toEqual(ok(token));
  });

  it('refuses an expired token in ASKMARCEL_TOKEN_ELEVATED with env_token_invalid, naming the variable and never its value, and does not fall back to the helper', async () => {
    const token = jwt({ exp: Math.floor(Date.now() / 1000) - 60, aud: GRAPH });
    const helper = helperFake();
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_ELEVATED: token }, helper).graphToken('elevated'));
    expect(message).toContain('ASKMARCEL_TOKEN_ELEVATED');
    expect(message).toContain('expired');
    expect(message).not.toContain(token);
    expect(helper.asked).toEqual([]);
  });

  it('refuses a value in ASKMARCEL_TOKEN_BASIC that is not a JWT at all', async () => {
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: 'not-a-token' }, helperFake()).graphToken('basic'));
    expect(message).toContain('not a JWT');
    expect(message).not.toContain('not-a-token');
  });

  it('refuses a token that decodes but holds a line break or a space, which no HTTP header can carry, and never echoes it', async () => {
    const graph = jwt({ exp: inAnHour(), aud: GRAPH });
    const [header, ...rest] = graph.split('.');
    const cases: ReadonlyArray<[Record<string, string>, 'basic' | 'ic3']> = [
      [{ ASKMARCEL_TOKEN_BASIC: `${graph}SECRETSIG\nPART2` }, 'basic'],
      [{ ASKMARCEL_TOKEN_BASIC: [`eyJ ${header?.slice(3) ?? ''}`, ...rest].join('.') }, 'basic'],
      [{ ASKMARCEL_TOKEN_IC3: `${jwt({ exp: inAnHour(), aud: IC3 })}\r\nX-Injected: 1` }, 'ic3'],
    ];
    for (const [env, tier] of cases) {
      const source = createEnvTokenSource(env, helperFake());
      const message = invalid(tier === 'basic' ? await source.graphToken(tier) : await source.substrateToken(tier));
      expect(message).toContain('it is not a JWT');
      expect(message).not.toContain('SECRETSIG');
      expect(message).not.toContain('Injected');
    }
  });

  it('refuses a Teams chat token set as the basic Graph token, because Graph would refuse its audience', async () => {
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: jwt({ exp: inAnHour(), aud: CHATSVCAGG }) }, helperFake()).graphToken('basic'));
    expect(message).toContain('Microsoft Graph');
  });

  it('refuses a Teams chat token set as the elevated token, saying that it must be issued for Microsoft Graph', async () => {
    const token = jwt({ exp: inAnHour(), aud: CHATSVCAGG });
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_ELEVATED: token }, helperFake()).graphToken('elevated'));
    expect(message).toContain('ASKMARCEL_TOKEN_ELEVATED');
    expect(message).toContain('it is not issued for Microsoft Graph');
    expect(message).not.toContain(token);
  });

  it('asks the helper for a tier whose variable is not set, even when another tier has one', async () => {
    const helper = helperFake();
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: jwt({ exp: inAnHour(), aud: GRAPH }) }, helper);
    expect(await source.graphToken('elevated')).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual(['elevated']);
  });

  it('reads an empty variable as not set, so the helper answers', async () => {
    const helper = helperFake();
    expect(await createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: '' }, helper).graphToken('basic')).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual(['basic']);
  });

  it('takes the chatsvcagg token and its region from the environment, with no helper call', async () => {
    const token = jwt({ exp: inAnHour(), aud: CHATSVCAGG });
    const helper = helperFake();
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_CHATSVCAGG: token, ASKMARCEL_TEAMS_REGION: 'emea' }, helper);
    expect(await source.substrateRegion('chatsvcagg')).toEqual(ok('emea' as TeamsRegion));
    expect(await source.substrateToken('chatsvcagg')).toEqual(ok(token));
    expect(helper.asked).toEqual([]);
  });

  it('takes the ic3 token from ASKMARCEL_TOKEN_IC3 when its audience is the IC3 service', async () => {
    const token = jwt({ exp: inAnHour(), aud: IC3 });
    expect(await createEnvTokenSource({ ASKMARCEL_TOKEN_IC3: token }, helperFake()).substrateToken('ic3')).toEqual(ok(token));
  });

  it('refuses a chatsvcagg token set as the ic3 token', async () => {
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_IC3: jwt({ exp: inAnHour(), aud: CHATSVCAGG }) }, helperFake()).substrateToken('ic3'));
    expect(message).toContain('ASKMARCEL_TOKEN_IC3');
    expect(message).toContain(IC3);
  });

  it('refuses a chat token in the environment that is not a JWT, or that has expired, saying which', async () => {
    const cases: ReadonlyArray<[Record<string, string>, 'chatsvcagg' | 'ic3', string]> = [
      [{ ASKMARCEL_TOKEN_CHATSVCAGG: 'opaque' }, 'chatsvcagg', 'it is not a JWT'],
      [{ ASKMARCEL_TOKEN_CHATSVCAGG: `x${jwt({ exp: inAnHour(), aud: CHATSVCAGG })}` }, 'chatsvcagg', 'it is not a JWT'],
      [{ ASKMARCEL_TOKEN_IC3: jwt({ exp: Math.floor(Date.now() / 1000) + 60, aud: IC3 }) }, 'ic3', 'it has expired'],
    ];
    for (const [env, tier, reason] of cases) {
      expect(invalid(await createEnvTokenSource(env, helperFake()).substrateToken(tier))).toContain(reason);
    }
  });

  it('refuses a chat token from the environment that has no ASKMARCEL_TEAMS_REGION beside it, rather than asking the helper for one', async () => {
    const helper = helperFake();
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_IC3: jwt({ exp: inAnHour(), aud: IC3 }) }, helper).substrateRegion('ic3'));
    expect(message).toContain('ASKMARCEL_TEAMS_REGION');
    expect(message).toContain('not set');
    expect(helper.asked).toEqual([]);
  });

  it('refuses an ASKMARCEL_TEAMS_REGION that is not a region name, so it never steers a chat request to another path', async () => {
    const env = { ASKMARCEL_TOKEN_CHATSVCAGG: jwt({ exp: inAnHour(), aud: CHATSVCAGG }), ASKMARCEL_TEAMS_REGION: 'emea/../admin' };
    const message = invalid(await createEnvTokenSource(env, helperFake()).substrateRegion('chatsvcagg'));
    expect(message).toContain('not a region name');
    expect(message).not.toContain('admin');
  });

  it('takes the region from the helper when the tier has no token variable, even when ASKMARCEL_TEAMS_REGION is set', async () => {
    const helper = helperFake();
    const source = createEnvTokenSource({ ASKMARCEL_TEAMS_REGION: 'emea' }, helper);
    expect(await source.substrateRegion('ic3')).toEqual(ok('amer' as TeamsRegion));
    expect(await source.substrateToken('ic3')).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual(['ic3 region', 'ic3']);
  });

  it('surfaces a chat token from the environment that the service refused, and never replays it with a token from the helper', async () => {
    const token = jwt({ exp: inAnHour(), aud: CHATSVCAGG });
    const helper = helperFake();
    const message = invalid(await createEnvTokenSource({ ASKMARCEL_TOKEN_CHATSVCAGG: token }, helper).substrateToken('chatsvcagg', { rejected: token }));
    expect(message).toContain('refused the token in ASKMARCEL_TOKEN_CHATSVCAGG');
    expect(message).not.toContain(token);
    expect(helper.asked).toEqual([]);
  });

  it('lets the helper replay a refused chat token when the tier has no token variable', async () => {
    const helper = helperFake();
    const refused = jwt({ exp: inAnHour(), aud: IC3 });
    expect(await createEnvTokenSource({}, helper).substrateToken('ic3', { rejected: refused })).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual(['ic3 past a refused token']);
  });

  it('always asks the helper for a guest token: guest tokens have no environment form', async () => {
    const helper = helperFake();
    const tenant = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
    const source = createEnvTokenSource({ ASKMARCEL_TOKEN_BASIC: jwt({ exp: inAnHour(), aud: GRAPH }) }, helper);
    expect(await source.guestToken(tenant)).toEqual(ok(HELPER_TOKEN));
    expect(helper.asked).toEqual([`guest ${tenant}`]);
  });
});
