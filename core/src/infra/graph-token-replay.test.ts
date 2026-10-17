import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createReadGraph } from './read-graph.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * The 401 replay of the basic and guest tiers (package split, per-tier policy):
 * a request Graph refuses because its token is invalid or expired is sent once
 * more with a token the source mints past the refused one. Any other 401 (a
 * partner-tenant file read on a home token, `invalidAudienceUri`) is an answer,
 * not a dead token, and is surfaced as it is. The elevated tier never replays:
 * it has no refresh token, so only a browser could renew it.
 */

const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
const OLD = accessTokenUnsafe('old-token');
const NEW = accessTokenUnsafe('new-token');

type Asked = string[];

// A source that hands out OLD until asked past it, then NEW (or `replay`).
const sourceAsked = (asked: Asked, replay: Result<AccessToken, TokenError> = ok(NEW)): TokenSource => {
  const answer = (what: string, rejected: AccessToken | undefined): Result<AccessToken, TokenError> => {
    asked.push(rejected === undefined ? what : `${what} past ${rejected}`);
    return rejected === undefined ? ok(OLD) : replay;
  };
  return {
    graphToken: async (tier, options) => answer(tier, options?.rejected),
    guestToken: async (tenant, options) => answer(`guest ${tenant}`, options?.rejected),
    substrateToken: async () => ok(OLD),
    substrateRegion: async () => ok('emea' as TeamsRegion),
  };
};

const refusal = (code: string, innerCode?: string): Response =>
  Response.json(
    { error: { code, message: 'Lifetime validation failed, the token is expired.', ...(innerCode === undefined ? {} : { innerError: { code: innerCode } }) } },
    { status: 401 }
  );

// Answers OLD with `refused` and NEW with 200, recording each bearer it saw.
const fetchRefusingOld = (seen: string[], refused: Response): FetchFn => {
  const fetchFn: FetchFn = async (_url, init) => {
    const bearer = new Headers(init?.headers).get('authorization') ?? '';
    seen.push(bearer);
    return bearer === `Bearer ${OLD}` ? refused.clone() : Response.json({ id: 'answered' });
  };
  return fetchFn;
};

describe('a basic request whose token Graph refused', () => {
  it('replays once with a token past the refused one when Graph says InvalidAuthenticationToken', async () => {
    const asked: Asked = [];
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked(asked), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    expect(await graph.get('/me')).toEqual(ok({ id: 'answered' }));
    expect(asked).toEqual(['basic', `basic past ${OLD}`]);
    expect(seen).toEqual([`Bearer ${OLD}`, `Bearer ${NEW}`]);
  });

  it('replays once when Graph says TokenExpired', async () => {
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked([]), fetchRefusingOld(seen, refusal('TokenExpired')));
    expect(await graph.get('/me')).toEqual(ok({ id: 'answered' }));
    expect(seen).toHaveLength(2);
  });

  it('never replays invalidAudienceUri: a partner-tenant file refuses every home token, however fresh', async () => {
    const asked: Asked = [];
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked(asked), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken', 'invalidAudienceUri')));
    const result = await graph.get('/drives/d1/items/i1');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(401);
      expect(result.error.code).toBe('invalidAudienceUri');
    }
    expect(asked).toEqual(['basic']);
    expect(seen).toHaveLength(1);
  });

  it('never replays a 403, whatever code it carries: only a 401 says the token is refused', async () => {
    const seen: string[] = [];
    const forbidden = Response.json({ error: { code: 'InvalidAuthenticationToken', message: 'Forbidden.' } }, { status: 403 });
    const result = await createReadGraph(sourceAsked([]), fetchRefusingOld(seen, forbidden)).get('/me');
    expect(result.ok).toBe(false);
    expect(seen).toHaveLength(1);
  });

  it('surfaces a 401 that names neither code, such as one with no body, without a replay', async () => {
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked([]), fetchRefusingOld(seen, new Response(null, { status: 401 })));
    const result = await graph.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.status).toBe(401);
    expect(seen).toHaveLength(1);
  });

  it('surfaces the second 401 as it is: a request is replayed once, never more', async () => {
    const seen: string[] = [];
    const fetchFn: FetchFn = async (_url, init) => {
      seen.push(new Headers(init?.headers).get('authorization') ?? '');
      return refusal('InvalidAuthenticationToken');
    };
    const result = await createReadGraph(sourceAsked([]), fetchFn).get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('InvalidAuthenticationToken');
    expect(seen).toEqual([`Bearer ${OLD}`, `Bearer ${NEW}`]);
  });

  it('does not send the request twice when the source has no other token than the refused one', async () => {
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked([], ok(OLD)), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    const result = await graph.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('InvalidAuthenticationToken');
    expect(seen).toHaveLength(1);
  });

  it('reports why no newer token could be had, with its code, when the replay finds none', async () => {
    const unavailable: Result<AccessToken, TokenError> = err({ type: 'auth_failed', message: 'Run login.', code: 'not_authenticated' });
    const graph = createReadGraph(sourceAsked([], unavailable), fetchRefusingOld([], refusal('InvalidAuthenticationToken')));
    expect(await graph.get('/me')).toEqual(err({ type: 'auth_failed', message: 'Run login.', code: 'not_authenticated' }));
  });

  it('replays a file read the same way', async () => {
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked([]), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    expect(await graph.getBinary('/me/drive/items/i1/content')).toEqual(ok({ id: 'answered' }));
    expect(seen).toHaveLength(2);
  });

  it('replays a write the same way: Graph ran nothing it refused, so the replay cannot duplicate it', async () => {
    const seen: string[] = [];
    const graph = createWriteGraph(sourceAsked([]), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    expect(await graph.patch('/me/messages/m1', { subject: 'Hi' })).toEqual(ok({ id: 'answered' }));
    expect(seen).toEqual([`Bearer ${OLD}`, `Bearer ${NEW}`]);
  });
});

describe('a guest request whose token the partner tenant refused', () => {
  it('replays once with a guest token of the same tenant past the refused one', async () => {
    const asked: Asked = [];
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked(asked), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    expect(await graph.getGuest('/drives/d1/items/i1', TENANT)).toEqual(ok({ id: 'answered' }));
    expect(await graph.getBinaryGuest('/drives/d1/items/i1/content', TENANT)).toEqual(ok({ id: 'answered' }));
    expect(asked).toEqual([`guest ${TENANT}`, `guest ${TENANT} past ${OLD}`, `guest ${TENANT}`, `guest ${TENANT} past ${OLD}`]);
    expect(seen).toHaveLength(4);
  });
});

describe('an elevated request whose token Graph refused', () => {
  it('is never replayed: the elevated tier has no refresh token', async () => {
    const asked: Asked = [];
    const seen: string[] = [];
    const graph = createReadGraph(sourceAsked(asked), fetchRefusingOld(seen, refusal('InvalidAuthenticationToken')));
    expect((await graph.getElevated('/me/chats')).ok).toBe(false);
    expect((await graph.getBinaryElevated('/drives/d1/items/i1/versions/1/content')).ok).toBe(false);
    expect(asked).toEqual(['elevated', 'elevated']);
    expect(seen).toHaveLength(2);
  });
});
