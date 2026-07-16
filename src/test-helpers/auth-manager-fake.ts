import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { AuthManager } from '../infra/auth.ts';

/**
 * Hand-written fake for the `AuthManager` secondary port (atelier rule 13 — the
 * `mock` namespace of `bun:test` is banned). The twin of `fakeGraphClient`, and
 * for the same reason: `AuthManager` was faked by full inline literals in five
 * test files (29 of them), so every new method broke all of them at once. The
 * 2026-06-15 LESSONS entry called this out for `GraphClient` "or a similarly
 * widely-faked port" — this is that port, given the same treatment.
 *
 * Defaults are the quiet ones: a usable basic token, every secondary tier
 * cancelled. Pass `overrides` for the behaviour a test actually cares about:
 *
 *   fakeAuthManager({ getGuestAccessToken: async () => ok(token) })
 *
 * Adding a method to `AuthManager` now touches THIS FILE, not five others.
 */
export const fakeAuthManager = (overrides: Partial<AuthManager> = {}): AuthManager => ({
  getAccessToken: async () => ({ ok: true, value: accessTokenUnsafe('tok') }),
  getElevatedAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } }),
  getGuestAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } }),
  getChatsvcaggAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } }),
  getChatsvcaggRegion: async () => 'emea',
  getIc3AccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } }),
  logout: async () => ({ ok: true, value: undefined }),
  getLastElevatedOutcome: () => null,
  getLastChatsvcaggOutcome: () => null,
  ...overrides,
});
