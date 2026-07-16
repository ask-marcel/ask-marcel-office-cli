import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { GraphClient } from '../infra/graph-client.ts';

/**
 * Hand-written fake for the `GraphClient` secondary port (atelier rule 13 — the
 * `mock` namespace of `bun:test` is banned). Every method defaults to an
 * `ok`-empty Result; pass `overrides` for the behaviour a given test cares
 * about (e.g. `fakeGraphClient({ get: async () => ok(user) })`).
 *
 * Centralising the method list here means a new `GraphClient` method is added in
 * exactly one place instead of the ~30 duplicated inline fakes this replaces.
 */
export const fakeGraphClient = (overrides: Partial<GraphClient> = {}): GraphClient => ({
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
  post: async () => ok({}),
  patch: async () => ok({}),
  getBinary: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () =>
    ok({
      scopes: [],
      audience: undefined,
      expiresAt: undefined,
      expiresInSeconds: undefined,
      elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
      chatsvcagg: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
      ic3: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'automatic' },
    }),
  ...overrides,
});
