import type { GraphClient } from '../infra/graph-client.ts';
import { fakeReadGraph } from './read-graph-fake.ts';
import { fakeWriteGraph } from './write-graph-fake.ts';

/**
 * Hand-written fake for the single package's full `GraphClient` (atelier rule 13
 * — the `mock` namespace of `bun:test` is banned): the read graph fake and the
 * write graph fake on one object. Every method defaults to an `ok`-empty
 * Result; pass `overrides` for the behaviour a given test cares about (e.g.
 * `fakeGraphClient({ get: async () => ok(user) })`).
 *
 * A new `ReadGraph` or `WriteGraph` member is added in the fake of its graph
 * (read-graph-fake.ts, write-graph-fake.ts), or in both when both graphs have
 * it; this builder only composes them. A test of one read or write command can
 * take the narrower fake, `fakeReadGraph` or `fakeWriteGraph`.
 */
export const fakeGraphClient = (overrides: Partial<GraphClient> = {}): GraphClient => ({
  ...fakeReadGraph(),
  ...fakeWriteGraph(),
  ...overrides,
});
