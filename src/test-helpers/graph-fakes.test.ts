import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { createReadGraph } from '../infra/read-graph.ts';
import { createWriteGraph } from '../infra/write-graph.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { fakeGraphClient } from './graph-client-fake.ts';
import { fakeReadGraph } from './read-graph-fake.ts';
import { fakeWriteGraph } from './write-graph-fake.ts';

// The fakes stand in for the two graphs, so each has exactly its graph's members:
// a read command tested on the read fake cannot reach a write member there either.
const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const members = (graph: object): ReadonlyArray<string> => Object.keys(graph).toSorted((a, b) => a.localeCompare(b));

describe('the graph fakes', () => {
  it('give the read fake the read graph members, the write fake the write graph members, and the full fake both', () => {
    const read = members(createReadGraph(tokens));
    const write = members(createWriteGraph(tokens));
    expect(members(fakeReadGraph())).toEqual(read);
    expect(members(fakeWriteGraph())).toEqual(write);
    expect(members(fakeGraphClient())).toEqual([...new Set([...read, ...write])].toSorted((a, b) => a.localeCompare(b)));
  });
});
