import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { FetchMock } from '@ask-marcel/office-core/test-helpers/fetch-mock.ts';
import { installFetchMock } from '@ask-marcel/office-core/test-helpers/fetch-mock.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createWriteClient } from './write-client.ts';

// The library helper of the write package: the write graph, signed by the tier
// variables of the process, else by the token helper of the auth package. The
// variable is set for each case and given back after it.

const VARIABLE = 'ASKMARCEL_TOKEN_BASIC';
let saved: string | undefined;
let fetches: FetchMock | undefined;

beforeEach(() => {
  saved = process.env[VARIABLE];
});

afterEach(() => {
  fetches?.restore();
  if (saved === undefined) delete process.env[VARIABLE];
  else process.env[VARIABLE] = saved;
});

describe('the write client of a library caller', () => {
  it('is the write graph: it has the basic-tier members only', () => {
    expect(Object.keys(createWriteClient()).toSorted((a, b) => a.localeCompare(b))).toEqual(['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put']);
  });

  it('signs a Graph request with the token of the tier variable', async () => {
    const token = unsignedJwt({ aud: 'https://graph.microsoft.com', exp: Math.floor(Date.now() / 1000) + 3600 });
    process.env[VARIABLE] = token;
    fetches = installFetchMock([{ match: () => true, respond: () => new Response(JSON.stringify({ id: 'me' }), { headers: { 'content-type': 'application/json' } }) }]);
    expect(await createWriteClient().get('/me')).toEqual(ok({ id: 'me' }));
    const headers = fetches.calls[0]?.init?.headers as Record<string, string> | undefined;
    expect(headers?.['Authorization']).toBe(`Bearer ${token}`);
  });
});
