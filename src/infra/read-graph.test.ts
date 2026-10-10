import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import type { ReadOnlyPostPath } from './read-graph.ts';
import { createReadGraph } from './read-graph.ts';

/*
 * The read guarantee (package split, D9): the graph a read command gets can
 * POST only to the two endpoints that change nothing, and has no put, patch,
 * delete or upload session at all. The path type refuses any other POST at
 * compile time; these tests pin the runtime half, for a caller that casts.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

type Sent = { readonly method: string; readonly url: string };

const fetchSent = (sent: Sent[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    sent.push({ method: init?.method ?? 'GET', url });
    return Response.json({ value: [] });
  };
  return fetchFn;
};

// A path the type forbids, as a caller that casts would pass it.
const writePath = (path: string): ReadOnlyPostPath => path as ReadOnlyPostPath;

const WRITE_MEMBERS = ['patch', 'put', 'delete', 'createUploadSession'];

describe('the read graph', () => {
  it('POSTs a search and a free/busy query, the two endpoints that change nothing', async () => {
    const sent: Sent[] = [];
    const graph = createReadGraph(tokens, fetchSent(sent));
    expect(await graph.post('/search/query', { requests: [] })).toEqual(ok({ value: [] }));
    expect(await graph.post('/me/calendar/getSchedule', { schedules: [] })).toEqual(ok({ value: [] }));
    expect(sent).toEqual([
      { method: 'POST', url: 'https://graph.microsoft.com/v1.0/search/query' },
      { method: 'POST', url: 'https://graph.microsoft.com/v1.0/me/calendar/getSchedule' },
    ]);
  });

  it('refuses any other POST before a request is made, so a cast cannot turn a read into a write', async () => {
    const sent: Sent[] = [];
    const result = await createReadGraph(tokens, fetchSent(sent)).post(writePath('/me/sendMail'), { message: {} });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.code).toBe('write_refused');
      expect(result.error.message).toBe('POST /me/sendMail was not sent: this graph only reads, and a read command may POST only to /search/query and /me/calendar/getSchedule.');
    }
    expect(sent).toEqual([]);
  });

  it('has no member that writes', () => {
    const graph = createReadGraph(tokens, fetchSent([]));
    expect(Object.keys(graph).filter((key) => WRITE_MEMBERS.includes(key))).toEqual([]);
  });
});

type Request = {
  readonly method: string;
  readonly url: string;
  readonly bearer: string | null;
  readonly prefer: string | null;
  readonly body: unknown;
  readonly redirect: unknown;
};

// Records every request as Graph would see it.
const fetchRequests = (requests: Request[]): FetchFn => {
  const fetchFn: FetchFn = async (url, init) => {
    const headers = new Headers(init?.headers);
    requests.push({ method: init?.method ?? '', url, bearer: headers.get('authorization'), prefer: headers.get('prefer'), body: init?.body, redirect: init?.redirect });
    return Response.json({});
  };
  return fetchFn;
};

describe('what the read graph sends', () => {
  const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
  const tiered: TokenSource = { ...tokens, graphToken: async (tier) => ok(accessTokenUnsafe(`${tier}-token`)) };

  it('sends each JSON read as a GET signed for its tier, with the headers its caller adds', async () => {
    const requests: Request[] = [];
    const graph = createReadGraph(tiered, fetchRequests(requests));
    await graph.get('/me/messages/delta', { Prefer: 'odata.maxpagesize=5' });
    await graph.getElevated('/me/chats');
    await graph.getGuest('/drives/d1', TENANT);
    expect(requests.map((r) => [r.method, r.bearer, r.prefer])).toEqual([
      ['GET', 'Bearer basic-token', 'odata.maxpagesize=5'],
      ['GET', 'Bearer elevated-token', null],
      ['GET', 'Bearer guest-token', null],
    ]);
  });

  it('sends a query POST with the basic token and its body', async () => {
    const requests: Request[] = [];
    await createReadGraph(tiered, fetchRequests(requests)).post('/search/query', { requests: [] });
    expect(requests.map((r) => [r.method, r.bearer, r.body])).toEqual([['POST', 'Bearer basic-token', '{"requests":[]}']]);
  });

  it('reads a file without following its redirect, so the CDN address comes back rather than the bytes', async () => {
    const requests: Request[] = [];
    const graph = createReadGraph(tiered, fetchRequests(requests));
    await graph.getBinary('/me/drive/items/i1/content');
    await graph.getBinaryElevated('/drives/d1/items/i1/versions/2/content');
    await graph.getBinaryGuest('/drives/d1/items/i1/content', TENANT);
    expect(requests.map((r) => [r.method, r.bearer, r.redirect])).toEqual([
      ['GET', 'Bearer basic-token', 'manual'],
      ['GET', 'Bearer elevated-token', 'manual'],
      ['GET', 'Bearer guest-token', 'manual'],
    ]);
  });

  it('names the call and its tier when the network fails on a JSON read', async () => {
    const failing: FetchFn = async () => {
      throw new Error('fetch failed');
    };
    const graph = createReadGraph(tokens, failing);
    const messages = await Promise.all([graph.get('/me'), graph.getElevated('/me/chats'), graph.getGuest('/drives/d1', TENANT)]);
    expect(messages.map((r) => (r.ok ? '' : r.error.message))).toEqual([
      'fetch failed (GET /me) — transient; retry once before treating as permanent',
      'fetch failed (GET /me/chats (elevated)) — transient; retry once before treating as permanent',
      `fetch failed (GET /drives/d1 (guest ${TENANT})) — transient; retry once before treating as permanent`,
    ]);
  });
});
