import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import type { FetchFn, GraphClient } from './graph-client.ts';
import { createGraphClient } from './graph-client.ts';

const IMAGE = 'https://as-api.asm.skype.com/v1/objects/0-weu-d1-abc/views/imgo';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

type Seen = { url: string; auth: string | null };

const clientWith = (answers: ReadonlyArray<Response>, seen: Seen[] = [], tokens: string[] = ['ic3-first', 'ic3-fresh']): GraphClient => {
  let call = 0;
  let token = 0;
  const fetchFn: FetchFn = async (input, init) => {
    seen.push({ url: String(input), auth: new Headers(init?.headers).get('authorization') });
    const answer = answers[Math.min(call, answers.length - 1)] ?? new Response(null, { status: 500 });
    call += 1;
    return answer;
  };
  const auth = fakeAuthManager({ getIc3AccessToken: async () => ({ ok: true, value: accessTokenUnsafe(tokens[Math.min(token++, tokens.length - 1)] ?? '') }) });
  return createGraphClient(auth, fetchFn);
};

describe('reading a pasted chat image from the Teams media service', () => {
  it('fetches it with the chat-history (IC3) bearer and answers the bytes', async () => {
    const seen: Seen[] = [];
    const result = await clientWith([new Response(PNG, { status: 200, headers: { 'content-type': 'image/png' } })], seen).teamsChatMedia(IMAGE);
    expect(result).toEqual({ ok: true, value: { contentType: 'image/png', size: 4, base64: 'iVBORw==' } });
    expect(seen).toEqual([{ url: IMAGE, auth: 'Bearer ic3-first' }]);
  });

  it('retries once with a fresh token when the cached one is refused', async () => {
    const seen: Seen[] = [];
    const answers = [new Response(null, { status: 401 }), new Response(PNG, { status: 200, headers: { 'content-type': 'image/png' } })];
    const result = await clientWith(answers, seen).teamsChatMedia('https://ch-prod.asyncgw.teams.microsoft.com/v1/objects/0-weu-d1-abc/views/imgo');
    expect(result.ok).toBe(true);
    expect(seen.map((s) => s.auth)).toEqual(['Bearer ic3-first', 'Bearer ic3-fresh']);
  });

  it('never sends the bearer outside the media hosts, nor over plain http', async () => {
    for (const url of ['https://evil.example/v1/objects/x/views/imgo', IMAGE.replace('https:', 'http:'), 'https://asm.skype.com.evil.example/x', 'not a url']) {
      const seen: Seen[] = [];
      const result = await clientWith([new Response(PNG, { status: 200 })], seen).teamsChatMedia(url);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
      expect(seen).toEqual([]);
    }
  });

  it('names a refused read, and a network failure', async () => {
    const refused = await clientWith([
      new Response(JSON.stringify({ errorCode: 404, message: 'gone' }), { status: 404, headers: { 'content-type': 'application/json' } }),
    ]).teamsChatMedia(IMAGE);
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('api_error');
    const failing: FetchFn = async () => {
      throw new TypeError('fetch failed');
    };
    const down = await createGraphClient(fakeAuthManager({ getIc3AccessToken: async () => ({ ok: true, value: accessTokenUnsafe('t') }) }), failing).teamsChatMedia(IMAGE);
    expect(down.ok).toBe(false);
    if (!down.ok) expect(down.error.type).toBe('network_error');
  });
});
