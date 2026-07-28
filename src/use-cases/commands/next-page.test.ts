import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute, meta } from './next-page.ts';

const trackingGraph = (): { graph: GraphClient; readonly calls: { readonly via: 'basic' | 'elevated' | 'guest'; readonly path: string; readonly tenantId?: string }[] } => {
  const calls: { via: 'basic' | 'elevated' | 'guest'; path: string; tenantId?: string }[] = [];
  return {
    calls,
    graph: fakeGraphClient({
      get: async (path: string) => {
        calls.push({ via: 'basic', path });
        return ok({});
      },
      getElevated: async (path: string) => {
        calls.push({ via: 'elevated', path });
        return ok({});
      },
      getGuest: async (path: string, tenantId: string) => {
        calls.push({ via: 'guest', path, tenantId });
        return ok({});
      },
    }),
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

  it('routes /chats/{id}/members nextLinks to graph.get (basic token — list-chat-members reads members via ChatMember.Read, no longer elevated)', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/chats/19:abc/members?$skiptoken=Q' });
    expect(calls).toEqual([{ via: 'basic', path: '/chats/19:abc/members?$skiptoken=Q' }]);
  });

  it('routes /chats/{id} metadata nextLinks to graph.getElevated (get-chat still needs Chat.ReadBasic)', async () => {
    const { graph, calls } = trackingGraph();
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/chats/19:abc?$select=id&$skiptoken=Q' });
    expect(calls).toEqual([{ via: 'elevated', path: '/chats/19:abc?$select=id&$skiptoken=Q' }]);
  });

  it('rejects a URL that does not start with the Graph v1.0 prefix without contacting the graph client, naming the expected form', async () => {
    const { graph, calls } = trackingGraph();
    const result = await execute(graph, { url: 'https://example.com/something' });
    expect(calls).toHaveLength(0);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    // Pin the refine message so the user is told what shape is expected, not a
    // bare "Invalid input" (kills the empty-message / dropped-message-object mutants).
    expect(result.error.message).toContain('Microsoft Graph v1.0 URL');
  });

  it('--url description documents the nextLink cursor source, the loop-until-absent pattern, and deltaLink', () => {
    const urlOption = meta.options.find((o) => o.key === 'url');
    expect(urlOption?.description).toContain('Full Graph v1.0 URL');
    expect(urlOption?.description).toContain('Example:');
    expect(urlOption?.description).toContain('Loop:');
    expect(urlOption?.description).toContain('deltaLink');
  });

  it('routes a partner-tenant drive-listing cursor to graph.getGuest when --tenant-id is given', async () => {
    const { graph, calls } = trackingGraph();
    const tenantId = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/drives/b!x/items/01ABC/children?$skiptoken=P', tenantId });
    expect(calls).toEqual([{ via: 'guest', path: '/drives/b!x/items/01ABC/children?$skiptoken=P', tenantId }]);
  });

  it('signs the cursor with the guest token even for a path that would otherwise be basic (tenant-id wins)', async () => {
    const { graph, calls } = trackingGraph();
    const tenantId = '6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04';
    await execute(graph, { url: 'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=P', tenantId });
    expect(calls.map((c) => c.via)).toEqual(['guest']);
  });

  it('rejects a malformed --tenant-id at the boundary without contacting the graph client', async () => {
    const { graph, calls } = trackingGraph();
    const result = await execute(graph, { url: 'https://graph.microsoft.com/v1.0/drives/b!x/items/01ABC/children?$skiptoken=P', tenantId: 'not-a-guid' });
    expect(calls).toHaveLength(0);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.message).toContain('tenant GUID');
  });

  it('--tenant-id description documents the partner-tenant guest-token behaviour and is optional', () => {
    const tenantOption = meta.options.find((o) => o.key === 'tenantId');
    expect(tenantOption?.required).toBe(false);
    expect(tenantOption?.description).toContain('PARTNER tenant');
    expect(tenantOption?.description).toContain('resolve-drive-share-link');
    expect(tenantOption?.description).toContain('guest token');
    expect(tenantOption?.description).toContain('invalidAudienceUri');
    expect(tenantOption?.description).toContain('Omit it');
  });
});
