import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * The write graph on its own (core, package split phase 3): a large upload goes
 * through an upload session in chunks, and a delete answers with no value.
 */

const tokens: TokenSource = {
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const MiB = 1024 * 1024;
const UPLOAD_URL = 'https://contoso.sharepoint.com/upload/session-1';
const TEMP_FILE = '/me/drive/root:/.ask-marcel-temp/a.docx';

type Sent = { readonly method: string; readonly url: string; readonly range: string | null; readonly authorization: string | null };

// Graph answers the session request with `session`, and each chunk PUT with the
// next of `chunks`; a chunk answer that is an Error is thrown, as a network failure.
const uploadGraph = (sent: Sent[], session: Response, chunks: ReadonlyArray<Response | Error>, onDelete: () => Response = () => new Response(null, { status: 204 })): FetchFn => {
  const answers = [...chunks];
  const fetchFn: FetchFn = async (url, init) => {
    const headers = new Headers(init?.headers);
    const method = init?.method ?? 'GET';
    sent.push({ method, url, range: headers.get('content-range'), authorization: headers.get('authorization') });
    if (method === 'POST') return session.clone();
    if (method === 'DELETE') return onDelete();
    const answer = answers.shift() ?? new Response(null, { status: 202 });
    if (answer instanceof Error) throw answer;
    return answer;
  };
  return fetchFn;
};

const sessionAt = (uploadUrl: string): Response => Response.json({ uploadUrl });

describe('a file larger than 4 MiB', () => {
  it('is sent in one chunk to the upload URL, unsigned, and answers with the item Graph created', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt(UPLOAD_URL), [Response.json({ id: 'item-1' }, { status: 201 })]));
    expect(await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1))).toEqual(ok({ id: 'item-1' }));
    expect(sent[1]).toEqual({ method: 'PUT', url: UPLOAD_URL, range: `bytes 0-${4 * MiB}/${4 * MiB + 1}`, authorization: null });
  });

  it('is sent in 5 MiB chunks, each accepted until the last one completes the file', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt(UPLOAD_URL), [new Response(null, { status: 202 }), Response.json({ id: 'item-2' }, { status: 200 })]));
    expect(await graph.put(TEMP_FILE, new Uint8Array(6 * MiB))).toEqual(ok({ id: 'item-2' }));
    expect(sent.slice(1).map((s) => s.range)).toEqual([`bytes 0-${5 * MiB - 1}/${6 * MiB}`, `bytes ${5 * MiB}-${6 * MiB - 1}/${6 * MiB}`]);
  });

  it('cancels the upload session and names the byte where a chunk was refused', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt(UPLOAD_URL), [new Response(null, { status: 202 }), new Response(null, { status: 507 })]));
    expect(await graph.put(TEMP_FILE, new Uint8Array(6 * MiB))).toEqual(err({ type: 'api_error', status: 507, message: `chunk PUT failed at byte ${5 * MiB}` }));
    expect(sent.at(-1)).toEqual({ method: 'DELETE', url: UPLOAD_URL, range: null, authorization: null });
  });

  it('cancels the upload session when a chunk fails on the network, even if the cancel fails too', async () => {
    const sent: Sent[] = [];
    const cancelFails = (): Response => {
      throw new Error('socket hang up');
    };
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt(UPLOAD_URL), [new Error('fetch failed')], cancelFails));
    const result = await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1));
    expect(result).toEqual(err({ type: 'network_error', message: 'fetch failed (PUT chunk @ byte 0) — transient; retry once before treating as permanent' }));
    expect(sent.map((s) => s.method)).toEqual(['POST', 'PUT', 'DELETE']);
  });

  it('fails when every chunk was accepted and none completed the file', async () => {
    const graph = createWriteGraph(tokens, uploadGraph([], sessionAt(UPLOAD_URL), [new Response(null, { status: 202 })]));
    expect(await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1))).toEqual(err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' }));
  });

  it('sends no chunk to an upload URL that is not a URL', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt('not a url'), []));
    expect(await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1))).toEqual(err({ type: 'network_error', message: 'createUploadSession returned an invalid uploadUrl' }));
    expect(sent.map((s) => s.method)).toEqual(['POST']);
  });

  it('sends no chunk to an upload URL outside the Microsoft hosts', async () => {
    const sent: Sent[] = [];
    const graph = createWriteGraph(tokens, uploadGraph(sent, sessionAt('https://upload.example.com/session'), []));
    expect(await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1))).toEqual(
      err({ type: 'network_error', message: 'uploadUrl host upload.example.com not in Microsoft allow-list' })
    );
    expect(sent.map((s) => s.method)).toEqual(['POST']);
  });

  it('reports a refused upload session as Graph answered it, and sends no chunk', async () => {
    const sent: Sent[] = [];
    const refused = Response.json({ error: { code: 'quotaLimitReached', message: 'The drive is full.' } }, { status: 507 });
    const graph = createWriteGraph(tokens, uploadGraph(sent, refused, []));
    expect(await graph.put(TEMP_FILE, new Uint8Array(4 * MiB + 1))).toEqual(
      err({ type: 'api_error', status: 507, message: 'quotaLimitReached: The drive is full.', code: 'quotaLimitReached' })
    );
    expect(sent.map((s) => s.method)).toEqual(['POST']);
  });
});

describe('a delete', () => {
  const deleting = (answer: () => Response): FetchFn => {
    const fetchFn: FetchFn = async () => answer();
    return fetchFn;
  };

  it('answers with no value when Graph removed the item', async () => {
    expect(
      await createWriteGraph(
        tokens,
        deleting(() => new Response(null, { status: 204 }))
      ).delete('/me/drive/items/temp-1')
    ).toEqual(ok(undefined));
  });

  it('reports the Graph error when the item could not be removed', async () => {
    const graph = createWriteGraph(
      tokens,
      deleting(() => Response.json({ error: { code: 'itemNotFound', message: 'The item is gone.' } }, { status: 404 }))
    );
    expect(await graph.delete('/me/drive/items/temp-1')).toEqual(err({ type: 'api_error', status: 404, message: 'itemNotFound: The item is gone.', code: 'itemNotFound' }));
  });

  it('names the delete and the request deadline when the network fails', async () => {
    const graph = createWriteGraph(
      tokens,
      deleting(() => {
        throw new DOMException('timed out', 'TimeoutError');
      })
    );
    expect(await graph.delete('/me/drive/items/temp-1')).toEqual(
      err({ type: 'network_error', message: 'request timed out after 60s (DELETE /me/drive/items/temp-1) — transient; retry once before treating as permanent' })
    );
  });

  it('sends nothing when no token can be had, and says why', async () => {
    const sent: string[] = [];
    const noToken: TokenSource = { ...tokens, graphToken: async () => err({ type: 'auth_failed', message: 'Sign in first.', code: 'not_authenticated' }) };
    const fetchFn: FetchFn = async (url) => {
      sent.push(url);
      return new Response(null, { status: 204 });
    };
    const graph = createWriteGraph(noToken, fetchFn);
    expect(await graph.delete('/me/drive/items/temp-1')).toEqual(err({ type: 'auth_failed', message: 'Sign in first.', code: 'not_authenticated' }));
    expect(await graph.put(TEMP_FILE, new Uint8Array(1))).toEqual(err({ type: 'auth_failed', message: 'Sign in first.', code: 'not_authenticated' }));
    expect(sent).toEqual([]);
  });
});

describe('a file the write graph reads', () => {
  it('comes back as bytes signed with the basic token', async () => {
    const signed: Array<string | null> = [];
    const fetchFn: FetchFn = async (_url, init) => {
      signed.push(new Headers(init?.headers).get('authorization'));
      return new Response(new Uint8Array([1]), { headers: { 'content-type': 'application/pdf' } });
    };
    expect(await createWriteGraph(tokens, fetchFn).getBinary('/me/messages/m/attachments/a/$value')).toEqual(ok({ contentType: 'application/pdf', size: 1, base64: 'AQ==' }));
    expect(signed).toEqual(['Bearer basic-token']);
  });
});
