import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { buildCommand, buildElevatedCommand } from './build-command.ts';

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
