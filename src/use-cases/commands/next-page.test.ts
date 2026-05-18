import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { execute } from './next-page.ts';

const trackingGraph = (): { graph: GraphClient; readonly calls: { readonly via: 'basic' | 'elevated'; readonly path: string }[] } => {
  const calls: { via: 'basic' | 'elevated'; path: string }[] = [];
  return {
    calls,
    graph: {
      get: async (path: string) => {
        calls.push({ via: 'basic', path });
        return ok({});
      },
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async (path: string) => {
        calls.push({ via: 'elevated', path });
        return ok({});
      },
      getBinaryElevated: async () => ok({}),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
      getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined }),
    },
  };
};

describe('next-page', () => {
  it('routes /me/messages nextLinks to graph.get (basic token)', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=ABC' });
    expect(calls).toEqual([{ via: 'basic', path: '/me/messages?$skiptoken=ABC' }]);
  });

  it('routes /me/chats nextLinks to graph.get (basic token) — chat commands no longer require elevation', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=XYZ' });
    expect(calls).toEqual([{ via: 'basic', path: '/me/chats?$skiptoken=XYZ' }]);
  });

  it('routes /chats/{id}/members nextLinks to graph.get (basic token) — matches list-chat-members which uses the regular token', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/chats/19:abc/members?$skiptoken=Q' });
    expect(calls).toEqual([{ via: 'basic', path: '/chats/19:abc/members?$skiptoken=Q' }]);
  });

  it('rejects a URL that does not start with the Graph v1.0 prefix without contacting the graph client', async () => {
    const { graph, calls } = trackingGraph();
    const result = await execute(graph, { url: 'https://example.com/something' });
    expect(calls).toHaveLength(0);
    expect(result.ok).toBe(false);
  });
});
