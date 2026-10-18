import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { TeamsRegion } from '@ask-marcel/office-core/domain/teams-region.ts';
import type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-client.ts';
import { createTokenSourceGraphClient } from './graph-client.ts';

// A usable basic token; every other tier is cancelled.
const basicOnly: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('tok')),
  guestToken: async () => err({ type: 'auth_cancelled' }),
  substrateToken: async () => err({ type: 'auth_cancelled' }),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

// An attachment's `$value` streams from Graph itself (no CDN redirect), so the
// whole transfer runs on this request: a 20 MB forwarded mail on a slow link
// outlived the 60-second budget of a JSON read (2026-09-27).
const hanging: FetchFn = async () => {
  throw Object.assign(new Error('The operation timed out.'), { name: 'TimeoutError' });
};

describe('a byte read straight from Graph', () => {
  it('runs on the five-minute transfer budget, and says so when it runs out', async () => {
    let signal: AbortSignal | undefined;
    const watching: FetchFn = async (_input, init) => {
      signal = init?.signal ?? undefined;
      return new Response(new Uint8Array([1, 2, 3]), { status: 200, headers: { 'content-type': 'application/octet-stream' } });
    };
    await createTokenSourceGraphClient(basicOnly, watching).getBinary('/me/messages/m1/attachments/a1/$value');
    expect(signal).toBeDefined();
    const result = await createTokenSourceGraphClient(basicOnly, hanging).getBinary('/me/messages/m1/attachments/a1/$value');
    expect(result.ok).toBe(false);
    if (!result.ok)
      expect(result.error.message).toBe('request timed out after 5min (GET /me/messages/m1/attachments/a1/$value (binary)) — transient; retry once before treating as permanent');
  });
});
