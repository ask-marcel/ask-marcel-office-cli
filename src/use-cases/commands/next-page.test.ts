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
      teamsChat: async () => ok({}),
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

  it('routes /me/chats nextLinks to graph.getElevated (round-8: chat commands re-elevated)', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=XYZ' });
    expect(calls).toEqual([{ via: 'elevated', path: '/me/chats?$skiptoken=XYZ' }]);
  });

  it('routes /chats/{id}/members nextLinks to graph.getElevated (round-8: list-chat-members re-elevated)', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/chats/19:abc/members?$skiptoken=Q' });
    expect(calls).toEqual([{ via: 'elevated', path: '/chats/19:abc/members?$skiptoken=Q' }]);
  });

  it('rejects a URL that does not start with the Graph v1.0 prefix without contacting the graph client', async () => {
    const { graph, calls } = trackingGraph();
    const result = await execute(graph, { url: 'https://example.com/something' });
    expect(calls).toHaveLength(0);
    expect(result.ok).toBe(false);
  });
});
