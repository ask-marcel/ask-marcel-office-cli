import { ok } from '../domain/result.ts';
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
      elevated: { available: false, expiresInSeconds: undefined },
      chatsvcagg: { available: false, expiresInSeconds: undefined },
      ic3: { available: false, expiresInSeconds: undefined },
    }),
  ...overrides,
});
