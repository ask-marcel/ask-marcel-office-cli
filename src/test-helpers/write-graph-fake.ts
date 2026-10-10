import { ok } from '../domain/result.ts';
import type { WriteGraph } from '../infra/write-graph.ts';

/**
 * Hand-written fake for the `WriteGraph` secondary port, the graph a write
 * command runs on (atelier rule 13: no `mock`). Every member defaults to an
 * `ok`-empty Result; pass `overrides` for the behaviour a test cares about
 * (e.g. `fakeWriteGraph({ post: async () => ok(draft) })`). Like the port, it
 * holds the basic tier only.
 */
export const fakeWriteGraph = (overrides: Partial<WriteGraph> = {}): WriteGraph => ({
  get: async () => ok({}),
  getBinary: async () => ok({}),
  fetchUrl: async () => ok({}),
  post: async () => ok({}),
  patch: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  ...overrides,
});
