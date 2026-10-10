import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import type { ReadGraph, ReadOnlyPostPath } from './read-graph.ts';
import { readGraphOf } from '../use-cases/commands/command-graph.ts';
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

describe('the read view of a graph that can also write', () => {
  // The single package's full client: its POST would send anything.
  const fullGraph = (posted: string[]): ReadGraph & Record<string, unknown> => ({
    ...createReadGraph(tokens, fetchSent([])),
    post: async (path: string) => {
      posted.push(path);
      return ok({});
    },
    patch: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
  });

  it('keeps only the read members, so a read command never holds a write method', () => {
    expect(Object.keys(readGraphOf(fullGraph([]))).toSorted((a, b) => a.localeCompare(b))).toEqual(
      Object.keys(createReadGraph(tokens, fetchSent([]))).toSorted((a, b) => a.localeCompare(b))
    );
  });

  it("sends each read to the graph's own member of that name, on a caller's graph whose members are methods that read `this`", async () => {
    const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
    const graph = {
      calls: [] as string[],
      async get(path: string) {
        this.calls.push(`get ${path}`);
        return ok({});
      },
      async getElevated(path: string) {
        this.calls.push(`getElevated ${path}`);
        return ok({});
      },
      async getGuest(path: string) {
        this.calls.push(`getGuest ${path}`);
        return ok({});
      },
      async getBinaryGuest(path: string) {
        this.calls.push(`getBinaryGuest ${path}`);
        return ok({});
      },
      async discoverTenantId(host: string) {
        this.calls.push(`discoverTenantId ${host}`);
        return ok(TENANT);
      },
      async teamsChat(path: string) {
        this.calls.push(`teamsChat ${path}`);
        return ok({});
      },
      async teamsChatIc3(path: string) {
        this.calls.push(`teamsChatIc3 ${path}`);
        return ok({});
      },
      async teamsChatMedia(url: string) {
        this.calls.push(`teamsChatMedia ${url}`);
        return ok({});
      },
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({});
      },
      async getBinary(path: string) {
        this.calls.push(`getBinary ${path}`);
        return ok({});
      },
      async getBinaryElevated(path: string) {
        this.calls.push(`getBinaryElevated ${path}`);
        return ok({});
      },
      async fetchUrl(url: string) {
        this.calls.push(`fetchUrl ${url}`);
        return ok({});
      },
    };
    const view = readGraphOf(graph);
    await view.get('/1');
    await view.getElevated('/2');
    await view.getGuest('/3', TENANT);
    await view.getBinaryGuest('/4', TENANT);
    await view.discoverTenantId('contoso.sharepoint.com');
    await view.teamsChat('/6');
    await view.teamsChatIc3('/7');
    await view.teamsChatMedia('https://8');
    await view.post('/search/query', {});
    await view.getBinary('/10');
    await view.getBinaryElevated('/11');
    await view.fetchUrl('https://12');
    expect(graph.calls).toEqual([
      'get /1',
      'getElevated /2',
      'getGuest /3',
      'getBinaryGuest /4',
      'discoverTenantId contoso.sharepoint.com',
      'teamsChat /6',
      'teamsChatIc3 /7',
      'teamsChatMedia https://8',
      'post /search/query',
      'getBinary /10',
      'getBinaryElevated /11',
      'fetchUrl https://12',
    ]);
  });

  it('passes a read-only POST through and refuses any other, though the graph under it would send it', async () => {
    const posted: string[] = [];
    const view = readGraphOf(fullGraph(posted));
    expect(await view.post('/search/query', {})).toEqual(ok({}));
    expect((await view.post(writePath('/me/messages'), {})).ok).toBe(false);
    expect(posted).toEqual(['/search/query']);
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
