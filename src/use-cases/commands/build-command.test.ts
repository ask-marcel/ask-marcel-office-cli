import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { buildCommand, buildElevatedCommand, buildElevatedListCommand, buildListCommand } from './build-command.ts';

const fakeGraph: GraphClient = {
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ({ ok: true, value: {} }),
  getBinaryElevated: async () => ({ ok: true, value: {} }),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
};

describe('buildCommand', () => {
  it('returns err({ type: "validation_error" }) with the zod message when schema validation fails', async () => {
    const cmd = buildCommand((p) => `/items/${p.id}`, z.object({ id: z.string().min(1) }));
    const result = await cmd.execute(fakeGraph, {});
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      if (result.error.type === 'validation_error') expect(result.error.message).toBe('id: Invalid input: expected string, received undefined');
    }
  });

  it('calls graph.get with the constructed path on valid params', async () => {
    let captured = '';
    const graph: GraphClient = {
      get: async (path: string) => {
        captured = path;
        return ok({});
      },
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async () => ({ ok: true, value: {} }),
      getBinaryElevated: async () => ({ ok: true, value: {} }),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildCommand((p) => `/items/${p.id}`, z.object({ id: z.string() }));
    const result = await cmd.execute(graph, { id: '42' });
    expect(result).toEqual(ok({}));
    expect(captured).toBe('/items/42');
  });
});

describe('buildElevatedCommand', () => {
  it('returns err({ type: "validation_error" }) with the zod message when schema validation fails', async () => {
    const cmd = buildElevatedCommand((p) => `/chats/${p.id}`, z.object({ id: z.string().min(1) }));
    const result = await cmd.execute(fakeGraph, {});
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      if (result.error.type === 'validation_error') expect(result.error.message).toBe('id: Invalid input: expected string, received undefined');
    }
  });

  it('calls graph.getElevated with the constructed path on valid params', async () => {
    let captured = '';
    const graph: GraphClient = {
      get: async () => ok({}),
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async (path: string) => {
        captured = path;
        return ok({});
      },
      getBinaryElevated: async () => ({ ok: true, value: {} }),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildElevatedCommand((p) => `/chats/${p.id}`, z.object({ id: z.string() }));
    const result = await cmd.execute(graph, { id: '19:abc' });
    expect(result).toEqual(ok({}));
    expect(captured).toBe('/chats/19:abc');
  });
});

describe('buildListCommand', () => {
  it('appends $top to the constructed path when --top is supplied', async () => {
    let captured = '';
    const graph: GraphClient = {
      get: async (path: string) => {
        captured = path;
        return ok({});
      },
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async () => ok({}),
      getBinaryElevated: async () => ok({}),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildListCommand(() => '/me/messages', z.object({}));
    await cmd.execute(graph, { top: '5' });
    expect(captured).toBe('/me/messages?$top=5');
  });

  it('rejects a non-numeric --top via Zod before reaching graph', async () => {
    let called = false;
    const graph: GraphClient = {
      get: async () => {
        called = true;
        return ok({});
      },
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async () => ok({}),
      getBinaryElevated: async () => ok({}),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildListCommand(() => '/me/messages', z.object({}));
    const result = await cmd.execute(graph, { top: 'lots' });
    expect(called).toBe(false);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('still validates the user-supplied schema (e.g. required IDs)', async () => {
    const cmd = buildListCommand((p) => `/sites/${p.siteId}`, z.object({ siteId: z.string().min(1) }));
    const result = await cmd.execute(fakeGraph, {});
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('exposes the merged schema (user fields + OData fragment) on the returned command', () => {
    const cmd = buildListCommand((p) => `/sites/${p.siteId}`, z.object({ siteId: z.string() }));
    const shape = (cmd.schema as unknown as { shape: Record<string, unknown> }).shape;
    expect(Object.keys(shape).toSorted()).toEqual(['expand', 'filter', 'orderby', 'select', 'siteId', 'skip', 'top']);
  });
});

describe('buildElevatedListCommand', () => {
  it('routes to graph.getElevated and applies OData passthrough', async () => {
    let captured = '';
    const graph: GraphClient = {
      get: async () => ok({}),
      post: async () => ok({}),
      getBinary: async () => ok({}),
      getElevated: async (path: string) => {
        captured = path;
        return ok({});
      },
      getBinaryElevated: async () => ok({}),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildElevatedListCommand(() => '/me/chats', z.object({}));
    await cmd.execute(graph, { top: '3', filter: "topic eq 'project'" });
    expect(captured).toBe("/me/chats?$top=3&$filter=topic%20eq%20'project'");
  });
});
