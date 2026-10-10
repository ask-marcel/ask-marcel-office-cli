import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { ReadGraph } from '../infra/read-graph.ts';

/**
 * Hand-written fake for the `ReadGraph` secondary port, the graph a read
 * command runs on (atelier rule 13: no `mock`). Every member defaults to an
 * `ok`-empty Result; pass `overrides` for the behaviour a test cares about
 * (e.g. `fakeReadGraph({ get: async () => ok(user) })`). It has no member that
 * writes, as the port has none.
 */
export const fakeReadGraph = (overrides: Partial<ReadGraph> = {}): ReadGraph => ({
  get: async () => ok({}),
  getElevated: async () => ok({}),
  getGuest: async () => ok({}),
  getBinaryGuest: async () => ok({}),
  // A placeholder tenant, never a real one: the repo's history was rewritten once
  // to purge real tenant identifiers, and a fixture is the easiest place to
  // reintroduce one.
  discoverTenantId: async () => ok(tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04')),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  teamsChatMedia: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  ...overrides,
});
