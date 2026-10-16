import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn } from './graph-request.ts';
import type { ReadGraph } from './read-graph.ts';
import { createReadGraph } from './read-graph.ts';

/*
 * How the Graph request core reads the bytes of a file, through the read graph
 * (core, package split phase 3): inline as JSON, text or base64, behind a
 * redirect, and through the CDN follow.
 */

const PARTNER = tenantIdUnsafe('72f988bf-86f1-41af-91ab-2d7cd011db47');

const tokens: TokenSource = {
  graphToken: async (tier) => ok(accessTokenUnsafe(`${tier}-token`)),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async () => ok(accessTokenUnsafe('substrate-token')),
  substrateRegion: async () => ok('emea' as TeamsRegion),
};

const answering = (answer: () => Response): FetchFn => {
  const fetchFn: FetchFn = async () => answer();
  return fetchFn;
};

const failing = (thrown: unknown): FetchFn => {
  const fetchFn: FetchFn = async () => {
    throw thrown;
  };
  return fetchFn;
};

const graphError = (status: number, error: unknown, headers: Record<string, string> = {}): Response => Response.json({ error }, { status, headers });

describe('the bytes of a file', () => {
  const binaryAnswering = (answer: () => Response): ReadGraph => createReadGraph(tokens, answering(answer));

  it('gives the CDN address of a redirect instead of the bytes', async () => {
    const redirect = (): Response => new Response(null, { status: 302, headers: { location: 'https://contoso.sharepoint.com/file.pdf' } });
    expect(await binaryAnswering(redirect).getBinary('/me/drive/items/i/content')).toEqual(ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/file.pdf' }));
  });

  it('reports a redirect with no address as the answer Graph gave', async () => {
    expect(await binaryAnswering(() => new Response(null, { status: 302 })).getBinary('/me/drive/items/i/content')).toEqual(
      err(expect.objectContaining({ type: 'api_error', status: 302 }))
    );
  });

  it('reads JSON as JSON, text with its size in UTF-8 bytes, and anything else as base64', async () => {
    expect(await binaryAnswering(() => Response.json({ value: 1 })).getBinary('/x')).toEqual(ok({ value: 1 }));
    expect(await binaryAnswering(() => new Response('héllo', { headers: { 'content-type': 'text/plain; charset=utf-8' } })).getBinary('/x')).toEqual(
      ok({ contentType: 'text/plain; charset=utf-8', size: 6, text: 'héllo' })
    );
    expect(await binaryAnswering(() => new Response(new Uint8Array([0, 255, 104]), { headers: { 'content-type': 'image/png' } })).getBinary('/x')).toEqual(
      ok({ contentType: 'image/png', size: 3, base64: 'AP9o' })
    );
  });

  it('reads XML and JavaScript as text, whatever the case of the content type', async () => {
    for (const contentType of ['application/atom+xml', 'APPLICATION/XML', 'application/javascript', 'Text/HTML']) {
      expect(await binaryAnswering(() => new Response('<a/>', { headers: { 'content-type': contentType } })).getBinary('/x')).toEqual(ok({ contentType, size: 4, text: '<a/>' }));
    }
  });

  it('names an answer with no content type as plain bytes', async () => {
    const untyped = (): Response => {
      const response = new Response(new Uint8Array([1, 2]));
      response.headers.delete('content-type');
      return response;
    };
    expect(await binaryAnswering(untyped).getBinary('/x')).toEqual(ok({ contentType: 'application/octet-stream', size: 2, base64: 'AQI=' }));
  });

  it('reports a refused file read as the Graph error', async () => {
    expect(await binaryAnswering(() => graphError(404, { code: 'itemNotFound', message: 'Gone.' })).getBinary('/x')).toEqual(
      err({ type: 'api_error', status: 404, message: 'itemNotFound: Gone.', code: 'itemNotFound' })
    );
  });

  it('signs an elevated and a guest file read with their own tokens', async () => {
    const signed: string[] = [];
    const recording: FetchFn = async (_url, init) => {
      signed.push(new Headers(init?.headers).get('authorization') ?? '');
      return Response.json({});
    };
    const graph = createReadGraph(tokens, recording);
    await graph.getBinaryElevated('/x');
    await graph.getBinaryGuest('/x', PARTNER);
    expect(signed).toEqual(['Bearer elevated-token', 'Bearer guest-token']);
  });
});

describe('the CDN follow', () => {
  it('reads a Microsoft CDN URL with no bearer', async () => {
    const sent: Array<string | null> = [];
    const recording: FetchFn = async (_url, init) => {
      sent.push(new Headers(init?.headers).get('authorization'));
      return new Response('page', { headers: { 'content-type': 'text/html' } });
    };
    expect(await createReadGraph(tokens, recording).fetchUrl('https://contoso.sharepoint.com/f.html')).toEqual(ok({ contentType: 'text/html', size: 4, text: 'page' }));
    expect(sent).toEqual([null]);
  });

  it('refuses a text that is not a URL, and a host outside the Microsoft list, before any request', async () => {
    const graph = createReadGraph(tokens, failing(new Error('must not be called')));
    expect(await graph.fetchUrl('not a url')).toEqual(err({ type: 'network_error', message: 'fetchUrl rejected: invalid URL not a url' }));
    expect(await graph.fetchUrl('https://contoso.sharepoint.com.evil.example/f')).toEqual(
      err({ type: 'network_error', message: 'fetchUrl rejected: host contoso.sharepoint.com.evil.example not in Microsoft allow-list' })
    );
  });

  it('accepts each Microsoft host of the list', async () => {
    const graph = createReadGraph(
      tokens,
      answering(() => Response.json({ ok: true }))
    );
    for (const url of [
      'https://a.onedrive.com/f',
      'https://a.live.com/f',
      'https://b.officeapps.live.com/f',
      'https://a.1drv.com/f',
      'https://graph.microsoft.com/f',
      'https://a.svc.ms/f',
    ]) {
      expect(await graph.fetchUrl(url)).toEqual(ok({ ok: true }));
    }
  });

  it('reports a refused CDN read and a failed one with the URL', async () => {
    expect(
      await createReadGraph(
        tokens,
        answering(() => new Response(null, { status: 403, statusText: 'Forbidden' }))
      ).fetchUrl('https://contoso.sharepoint.com/f')
    ).toEqual(err({ type: 'api_error', status: 403, message: 'Forbidden' }));
    expect(await createReadGraph(tokens, failing(new Error('fetch failed'))).fetchUrl('https://contoso.sharepoint.com/f')).toEqual(
      err({ type: 'network_error', message: 'fetch failed (GET https://contoso.sharepoint.com/f (CDN follow)) — transient; retry once before treating as permanent' })
    );
  });
});
