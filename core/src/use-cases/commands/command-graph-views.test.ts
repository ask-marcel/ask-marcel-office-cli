import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import { ok } from '../../domain/result.ts';
import type { TeamsRegion } from '../../domain/teams-region.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import type { FetchFn } from '../../infra/graph-request.ts';
import type { ReadGraph, ReadOnlyPostPath } from '../../infra/read-graph.ts';
import { createReadGraph } from '../../infra/read-graph.ts';
import { createWriteGraph } from '../../infra/write-graph.ts';
import type { TokenSource } from '../ports/token-source.ts';
import { readGraphOf, writeGraphOf } from './command-graph.ts';

/*
 * The view of a graph each command gets (package split, D9). The registry
 * gives a read command the read view of the graph its caller passed, and a
 * write command the write view. These are pure functions over a graph; the
 * graphs under them are real, built on a fetch that answers every request.
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

// A Graph that answers every request with a fresh copy of `response`.
const answering = (response: Response): FetchFn => {
  const fetchFn: FetchFn = async () => response.clone();
  return fetchFn;
};

const writeGraphAnswering = (response: Response): ReturnType<typeof createWriteGraph> => createWriteGraph(tokens, answering(response));

const WRITE_GRAPH_MEMBERS = ['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put'];

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

describe('the write view of a graph that can also read every tier', () => {
  it('keeps only the write graph members, so a write command never holds an elevated, guest or chat reader', () => {
    const full = { ...createReadGraph(tokens, answering(Response.json({}))), ...writeGraphAnswering(Response.json({})) };
    expect(Object.keys(writeGraphOf(full)).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_GRAPH_MEMBERS);
  });

  it("sends each write and read to the graph's own member of that name, on a caller's graph whose members are methods that read `this`", async () => {
    const graph = {
      calls: [] as string[],
      async get(path: string) {
        this.calls.push(`get ${path}`);
        return ok({});
      },
      async getBinary(path: string) {
        this.calls.push(`getBinary ${path}`);
        return ok({});
      },
      async fetchUrl(url: string) {
        this.calls.push(`fetchUrl ${url}`);
        return ok({});
      },
      async post(path: string) {
        this.calls.push(`post ${path}`);
        return ok({});
      },
      async patch(path: string) {
        this.calls.push(`patch ${path}`);
        return ok({});
      },
      async put(path: string) {
        this.calls.push(`put ${path}`);
        return ok({});
      },
      async delete(path: string) {
        this.calls.push(`delete ${path}`);
        return ok(undefined);
      },
    };
    const view = writeGraphOf(graph);
    await view.get('/1');
    await view.getBinary('/2');
    await view.fetchUrl('https://3');
    await view.post('/4', {});
    await view.patch('/5', {});
    await view.put('/6', new Uint8Array(1));
    await view.delete('/7');
    expect(graph.calls).toEqual(['get /1', 'getBinary /2', 'fetchUrl https://3', 'post /4', 'patch /5', 'put /6', 'delete /7']);
  });
});
