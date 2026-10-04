import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-client.ts';
import { createTokenSourceGraphClient } from './graph-client.ts';

type Seen = { url: string; auth: string | null };

// A token source whose two chat tiers live in different regions, as two
// answers of a token helper could, and that records which tier it was asked.
const sourceAsked = (asked: string[]): TokenSource => ({
  graphToken: async () => ok(accessTokenUnsafe('graph')),
  guestToken: async () => ok(accessTokenUnsafe('guest')),
  substrateToken: async (tier) => ok(accessTokenUnsafe(`${tier}-token`)),
  substrateRegion: async (tier) => {
    asked.push(tier);
    return ok((tier === 'ic3' ? 'amer' : 'apac') as TeamsRegion);
  },
});

// Records each request and answers it with an empty chat page.
const fetchSeen = (seen: Seen[]): FetchFn => {
  const fetchFn: FetchFn = async (input, init) => {
    seen.push({ url: String(input), auth: new Headers(init?.headers).get('authorization') });
    return new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } });
  };
  return fetchFn;
};

describe('a chat request signed by a token source', () => {
  it('asks for the region of the tier whose token signs the request, for each of the two chat services', async () => {
    const asked: string[] = [];
    const seen: Seen[] = [];
    const client = createTokenSourceGraphClient(sourceAsked(asked), fetchSeen(seen));
    await client.teamsChat('/v1/users/ME/conversations');
    await client.teamsChatIc3('/v1/users/ME/conversations');
    expect(asked).toEqual(['chatsvcagg', 'ic3']);
    expect(seen).toEqual([
      { url: 'https://teams.microsoft.com/api/csa/apac/v1/users/ME/conversations', auth: 'Bearer chatsvcagg-token' },
      { url: 'https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations', auth: 'Bearer ic3-token' },
    ]);
  });
});
