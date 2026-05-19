import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import type { AuthManager } from './auth.ts';
import type { FetchFn } from './graph-client.ts';
import { createGraphClient } from './graph-client.ts';

const fakeAuth = (): AuthManager => ({
  getAccessToken: async () => ok(accessTokenUnsafe('test-token')),
  getElevatedAccessToken: async () => ok(accessTokenUnsafe('test-elevated-token')),
  logout: async () => ok(undefined),
  getLastElevatedOutcome: () => null,
});

const fakeFetch = (responses: Array<{ match: (url: string) => boolean; body: unknown; status?: number }>): FetchFn => {
  return async (url: string) => {
    const handler = responses.find((r) => r.match(url));
    if (!handler) throw new Error(`no fetch handler matched ${url}`);
    return Response.json(handler.body, { status: handler.status ?? 200 });
  };
};

const timeoutFetch: FetchFn = async () => {
  const e = new Error('signal timed out');
  e.name = 'TimeoutError';
  throw e;
};

describe('graph client', () => {
  it('makes authenticated GET requests to the Graph API', async () => {
    const fetchFn = fakeFetch([{ match: (url) => url === 'https://graph.microsoft.com/v1.0/me/drives', body: { value: [{ id: 'drive-1', name: 'OneDrive' }] } }]);

    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/drives');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'drive-1', name: 'OneDrive' }] });
  });

  it('returns an error when the API returns an error status', async () => {
    const fetchFn = fakeFetch([{ match: (url) => url.includes('/me/drives'), body: { error: { message: 'not found' } }, status: 404 }]);

    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/drives');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('api_error');
  });

  it('returns an error when auth fails', async () => {
    const client = createGraphClient({
      getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed' as const, message: 'no auth' } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    });
    const result = await client.get('/me/drives');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('returns network_error when fetch throws', async () => {
    const throwingFetch: FetchFn = async () => {
      throw new Error('fetch failed');
    };

    const client = createGraphClient(fakeAuth(), throwingFetch);
    const result = await client.get('/me/drives');
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('network_error');
      expect(result.error.message).toContain('fetch failed');
      expect(result.error.message).toContain('GET /me/drives');
      expect(result.error.message).toContain('retry');
    }
  });

  it('truncates the granted-scope dump from `Missing scope permissions` 403 errors (audit round-8 Wave F)', async () => {
    const longScopeDump =
      "Forbidden: Missing scope permissions on the request. API requires one of 'Chat.ReadBasic, Chat.Read, Chat.ReadWrite'. Scopes on the request 'profile,openid,email,User.Read,Mail.Read,Calendars.Read,Tasks.Read,Sites.Read.All,Notes.Read.All,Group.Read.All,Team.ReadBasic.All,Channel.ReadBasic.All,People.Read,MailboxSettings.Read,Files.Read.All'";
    const fetchFn = fakeFetch([{ match: (url) => url.includes('/me/chats'), body: { error: { code: 'Forbidden', message: longScopeDump } }, status: 403 }]);

    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/chats');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain("API requires one of 'Chat.ReadBasic, Chat.Read, Chat.ReadWrite'");
      expect(result.error.message).not.toContain('User.Read,Mail.Read');
      expect(result.error.message).toContain('scopes-check');
    }
  });

  it('falls back to the response status text when the error body is not parseable JSON', async () => {
    const fetchFn: FetchFn = async () => new Response('not-json-at-all', { status: 503, statusText: 'Service Unavailable' });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(503);
      expect(result.error.message).toBe('Service Unavailable');
    }
  });

  it('reports an Auth cancelled message when the auth manager returns auth_cancelled', async () => {
    const cancelledAuth = {
      getAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(cancelledAuth);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('Auth cancelled');
    }
  });

  it('reports a generic network_error message when fetch throws a non-Error value', async () => {
    const throwingFetch: FetchFn = async () => {
      throw 'string thrown';
    };
    const client = createGraphClient(fakeAuth(), throwingFetch);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('string thrown');
    }
  });

  it('reports a generic network_error message when fetch throws a non-string non-Error value', async () => {
    const throwingFetch: FetchFn = async () => {
      throw { weird: 'thing' };
    };
    const client = createGraphClient(fakeAuth(), throwingFetch);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('network request failed');
    }
  });

  it('maps a TimeoutError thrown by AbortSignal.timeout to "request timed out after 60s" on Graph JSON GETs (the short-tier budget)', async () => {
    const client = createGraphClient(fakeAuth(), timeoutFetch);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('request timed out after 60s');
    }
  });

  // Audit v1.0.0 — SharePoint PDF download timeout fix. CDN body transfers
  // for multi-MB files take longer than the JSON-tier 60s budget, so
  // fetchUrl / simplePut / chunkedPut sit on a separate 5-minute tier. The
  // network-error message reflects which tier fired so a caller knows
  // which scale of failure they hit.
  it('maps a TimeoutError on fetchUrl (CDN body transfer) to "request timed out after 5min" — the long-tier budget for binary transfers', async () => {
    const client = createGraphClient(fakeAuth(), timeoutFetch);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/big.pdf');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('request timed out after 5min');
    }
  });

  it('maps a TimeoutError on the simple PUT path (≤4 MiB upload) to "request timed out after 5min" — the long-tier budget applies symmetrically to uploads', async () => {
    const client = createGraphClient(fakeAuth(), timeoutFetch);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/foo.bin', new Uint8Array(16));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('request timed out after 5min');
    }
  });

  it('maps an AbortError to "request aborted"', async () => {
    const abortFetch: FetchFn = async () => {
      const e = new Error('aborted');
      e.name = 'AbortError';
      throw e;
    };
    const client = createGraphClient(fakeAuth(), abortFetch);
    const result = await client.get('/me');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('request aborted');
    }
  });

  it('passes an AbortSignal.timeout signal to fetch (verifies the timeout is wired in)', async () => {
    let capturedSignal: AbortSignal | null | undefined = null;
    const captureFetch: FetchFn = async (_url, init) => {
      capturedSignal = init?.signal as AbortSignal | undefined;
      return Response.json({ ok: true });
    };
    const client = createGraphClient(fakeAuth(), captureFetch);
    await client.get('/me');
    expect(capturedSignal).toBeInstanceOf(AbortSignal);
  });

  it('fetchUrl follows a CDN URL and returns the parsed JSON when content-type advertises JSON', async () => {
    const fetchFn: FetchFn = async (url) => {
      expect(url).toBe('https://contoso.sharepoint.com/sites/x/file.json');
      return Response.json({ kind: 'json' });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/file.json');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ kind: 'json' });
  });

  it('fetchUrl returns text content as { contentType, size, text }', async () => {
    const fetchFn: FetchFn = async () => new Response('<html>hi</html>', { status: 200, headers: { 'content-type': 'text/html' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/page.html');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; text: string };
      expect(v.contentType).toBe('text/html');
      expect(v.text).toBe('<html>hi</html>');
    }
  });

  it('fetchUrl returns binary content base64-encoded for non-text/non-JSON responses', async () => {
    const bytes = new Uint8Array([0xff, 0xd8]);
    const fetchFn: FetchFn = async () => new Response(bytes, { status: 200, headers: { 'content-type': 'image/jpeg' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/photo.jpg');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; base64: string };
      expect(v.contentType).toBe('image/jpeg');
      expect(v.size).toBe(2);
    }
  });

  it('fetchUrl rejects URLs whose host is not on the Microsoft allow-list (Hardening #3)', async () => {
    const captureFetch: FetchFn = async () => {
      throw new Error('should not have been called');
    };
    const client = createGraphClient(fakeAuth(), captureFetch);
    const result = await client.fetchUrl('https://attacker.example.com/exfil');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('not in Microsoft allow-list');
    }
  });

  it('fetchUrl accepts the *.svc.ms Office Online conversion CDN that Graph hands out for ?format=html', async () => {
    const fetchFn: FetchFn = async () => new Response('<html>converted</html>', { status: 200, headers: { 'content-type': 'text/html' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://francecentral1-mediap.svc.ms/transform/htmlview?cs=fFNoYXJlUG9pbnQ');
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { text: string }).text).toBe('<html>converted</html>');
  });

  it('fetchUrl sends a permissive Accept header so the Office Online CDN does not return 406 Not Acceptable', async () => {
    let capturedAccept: string | null = null;
    const fetchFn: FetchFn = async (_url, init) => {
      const headers = new Headers(init?.headers);
      capturedAccept = headers.get('accept');
      return new Response('<html>x</html>', { status: 200, headers: { 'content-type': 'text/html' } });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    await client.fetchUrl('https://francecentral1-mediap.svc.ms/transform/htmlview');
    expect(capturedAccept).not.toBeNull();
    expect(capturedAccept ?? '').toContain('text/html');
    expect(capturedAccept ?? '').toContain('*/*');
  });

  it('fetchUrl rejects unparseable URL strings as a clear network_error', async () => {
    const client = createGraphClient(fakeAuth(), (async () => new Response()) as FetchFn);
    const result = await client.fetchUrl('not-a-url');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('invalid URL');
    }
  });

  it('fetchUrl reports api_error on non-2xx responses with status text', async () => {
    const fetchFn: FetchFn = async () => new Response('', { status: 503, statusText: 'Service Unavailable' });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/page.html');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(503);
      expect(result.error.message).toBe('Service Unavailable');
    }
  });

  it('fetchUrl extracts innererror.code from a JSON error body so the user sees the specific failure code, not just statusText', async () => {
    const body = JSON.stringify({
      error: {
        code: 'notSupported',
        message: 'An exception occurred while executing within the Sandbox',
        innererror: { code: 'Sandbox_InputFormatNotSupported' },
      },
    });
    const fetchFn: FetchFn = async () => new Response(body, { status: 406, statusText: 'Not Acceptable', headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://francecentral1-mediap.svc.ms/transform/html?x=1');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(406);
      expect(result.error.message).toBe('Sandbox_InputFormatNotSupported: An exception occurred while executing within the Sandbox');
    }
  });

  it('drops the empty-string `code` prefix Graph Planner returns on bogus IDs (was rendering as `": message"` — audit v1.0.0 §2.7)', async () => {
    const body = JSON.stringify({ error: { code: '', message: 'The requested item is not found.' } });
    const fetchFn: FetchFn = async () => new Response(body, { status: 404, statusText: 'Not Found', headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/planner/tasks/BOGUS');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toBe('The requested item is not found.');
      expect(result.error.message.startsWith(':')).toBe(false);
    }
  });

  it('rewrites an empty `UnknownError:` Graph body into an actionable retry hint (audit round-6 §1.2 — was leaking to the LLM with no recovery info)', async () => {
    const body = JSON.stringify({ error: { code: 'UnknownError', message: '' } });
    const fetchFn: FetchFn = async () => new Response(body, { status: 500, statusText: 'Internal Server Error', headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/messages/whatever');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('UnknownError:');
      expect(result.error.message).toContain('transient backend glitch');
      expect(result.error.message).toContain('retry once');
    }
  });

  it('also rewrites the whitespace-only `UnknownError: ` case (the form list-calendar-events-delta originally caught)', async () => {
    const body = JSON.stringify({ error: { code: 'UnknownError', message: ' ' } });
    const fetchFn: FetchFn = async () => new Response(body, { status: 500, headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/messages/whatever');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('transient backend glitch');
    }
  });

  it('leaves a NON-empty UnknownError body unchanged (only the empty/whitespace case is rewritten)', async () => {
    const body = JSON.stringify({ error: { code: 'UnknownError', message: 'specific Graph diagnostic' } });
    const fetchFn: FetchFn = async () => new Response(body, { status: 500, headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.get('/me/messages/whatever');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toBe('UnknownError: specific Graph diagnostic');
    }
  });

  it('fetchUrl also extracts the camelCase innerError.code (SharePoint streamContent uses camelCase, Graph uses lowercase)', async () => {
    const body = JSON.stringify({
      error: {
        code: 'accessDenied',
        innerError: { code: 'logicalPermissionAccessDenied' },
        message: 'The calling application is enrolled in logical permissions and is not permitted to call this API.',
      },
    });
    const fetchFn: FetchFn = async () => new Response(body, { status: 403, statusText: 'Forbidden', headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/_api/v2.0/drives/x/items/y/versions/3.0/streamContent');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(403);
      expect(result.error.message).toContain('logicalPermissionAccessDenied');
    }
  });

  it('fetchUrl wraps network errors via the same networkErrorMessage helper, with the CDN URL labeled', async () => {
    const throwing: FetchFn = async () => {
      throw new Error('socket reset');
    };
    const client = createGraphClient(fakeAuth(), throwing);
    const result = await client.fetchUrl('https://contoso.sharepoint.com/sites/x/page.html');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('socket reset');
      expect(result.error.message).toContain('CDN follow');
    }
  });

  it('getBinary returns the Location header as @microsoft.graph.downloadUrl on a 302 redirect', async () => {
    const fetchFn: FetchFn = async () => new Response(null, { status: 302, headers: { location: 'https://cdn.example/signed?token=abc' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/me/drive/items/i1/content');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@microsoft.graph.downloadUrl': 'https://cdn.example/signed?token=abc' });
  });

  it('getBinary returns a 3xx without a Location header as an api_error', async () => {
    const fetchFn: FetchFn = async () => new Response('weird', { status: 304 });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/me/photo/$value');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(304);
    }
  });

  it('getBinary returns text/* responses as a { text } envelope rather than base64', async () => {
    const html = '<html lang="en"><body>OneNote page</body></html>';
    const fetchFn: FetchFn = async () => new Response(html, { status: 200, headers: { 'content-type': 'text/html' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/me/onenote/pages/p1/content');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; text: string };
      expect(v.contentType).toBe('text/html');
      expect(v.text).toBe(html);
      expect(v.size).toBe(html.length);
    }
  });

  it('getBinary base64-encodes binary bodies and reports content-type and size', async () => {
    const bytes = new Uint8Array([0xff, 0xd8, 0xff, 0xe0]);
    const fetchFn: FetchFn = async () => new Response(bytes, { status: 200, headers: { 'content-type': 'image/jpeg' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/me/photo/$value');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; base64: string };
      expect(v.contentType).toBe('image/jpeg');
      expect(v.size).toBe(4);
      expect(v.base64).toBe(btoa(String.fromCharCode(0xff, 0xd8, 0xff, 0xe0)));
    }
  });

  it('getBinary returns parsed JSON when the response advertises application/json', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.context': 'foo', value: 'bar' });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/some/json/binary/endpoint');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ '@odata.context': 'foo', value: 'bar' });
  });

  it('getBinary returns api_error on a non-redirect non-OK response', async () => {
    const fetchFn: FetchFn = async () =>
      new Response(JSON.stringify({ error: { message: 'no permission' } }), {
        status: 403,
        headers: { 'content-type': 'application/json' },
      });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getBinary('/me/photo/$value');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(403);
      expect(result.error.message).toBe('no permission');
    }
  });

  it('getBinary surfaces auth_failed when the auth manager fails', async () => {
    const failingAuth: AuthManager = {
      getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'token gone' } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failingAuth);
    const result = await client.getBinary('/me/photo/$value');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('getBinary surfaces network_error when fetch throws, labeled with the binary path', async () => {
    const throwing: FetchFn = async () => {
      throw new Error('socket');
    };
    const client = createGraphClient(fakeAuth(), throwing);
    const result = await client.getBinary('/me/photo/$value');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('socket');
      expect(result.error.message).toContain('/me/photo/$value');
    }
  });

  it('makes authenticated POST requests with a JSON-serialised body', async () => {
    let captured: { url: string; method?: string; body?: string } | null = null;
    const fetchFn: FetchFn = async (url, init) => {
      captured = { url, method: init?.method, body: typeof init?.body === 'string' ? init.body : undefined };
      return Response.json({ value: [{ hits: [{ rank: 1 }] }] });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.post('/search/query', { requests: [{ entityTypes: ['chatMessage'] }] });
    expect(result.ok).toBe(true);
    expect(captured).not.toBeNull();
    if (captured !== null) {
      const c = captured as { url: string; method?: string; body?: string };
      expect(c.url).toBe('https://graph.microsoft.com/v1.0/search/query');
      expect(c.method).toBe('POST');
      expect(c.body).toBe(JSON.stringify({ requests: [{ entityTypes: ['chatMessage'] }] }));
    }
  });

  it('put with body ≤ 4 MiB takes the simple PUT path with :/content suffix', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      return Response.json({ id: 'i-new', name: 'small.bin' });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/abc', new Uint8Array(1024), 'application/octet-stream');
    expect(result.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe('PUT');
    expect(calls[0]?.url).toBe('https://graph.microsoft.com/v1.0/me/drive/root:/.ask-marcel-temp/abc:/content');
  });

  it('put with body 12 MiB takes the chunked-session path: createUploadSession + 3 chunk PUTs + final driveItem', async () => {
    const total = 12 * 1024 * 1024;
    const calls: Array<{ url: string; method?: string; range?: string }> = [];
    let chunkCount = 0;
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method, range: (init?.headers as Record<string, string> | undefined)?.['Content-Range'] });
      if (url.endsWith(':/createUploadSession')) {
        return Response.json({ uploadUrl: 'https://contoso-my.sharepoint.com/upload-session/abc' });
      }
      chunkCount += 1;
      if (chunkCount < 3) return new Response(JSON.stringify({ nextExpectedRanges: ['x-y'] }), { status: 202, headers: { 'content-type': 'application/json' } });
      return Response.json({ id: 'i-big', name: 'big.bin', size: total });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/big', new Uint8Array(total));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ id: 'i-big', name: 'big.bin', size: total });
    expect(calls).toHaveLength(4);
    expect(calls[0]?.url).toContain(':/createUploadSession');
    expect(calls[1]?.range).toBe(`bytes 0-${5 * 1024 * 1024 - 1}/${total}`);
    expect(calls[2]?.range).toBe(`bytes ${5 * 1024 * 1024}-${10 * 1024 * 1024 - 1}/${total}`);
    expect(calls[3]?.range).toBe(`bytes ${10 * 1024 * 1024}-${total - 1}/${total}`);
  });

  it('put surfaces api_error and DELETEs the upload session when a chunk PUT fails', async () => {
    const total = 6 * 1024 * 1024;
    const calls: Array<{ url: string; method?: string }> = [];
    let chunkCount = 0;
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith(':/createUploadSession')) {
        return Response.json({ uploadUrl: 'https://contoso-my.sharepoint.com/upload-session/x' });
      }
      if (init?.method === 'DELETE') return new Response(null, { status: 204 });
      chunkCount += 1;
      if (chunkCount === 1) return new Response('boom', { status: 500 });
      return Response.json({ id: 'should-not-finish' });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(500);
      expect(result.error.message).toContain('chunk PUT failed at byte 0');
    }
    expect(calls.some((c) => c.method === 'DELETE')).toBe(true);
  });

  it('put surfaces err from createUploadSession itself without attempting any chunk PUT', async () => {
    const total = 6 * 1024 * 1024;
    let chunkAttempts = 0;
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith(':/createUploadSession')) {
        return new Response(JSON.stringify({ error: { message: 'no permission' } }), { status: 403, headers: { 'content-type': 'application/json' } });
      }
      if (init?.method === 'PUT') chunkAttempts += 1;
      return Response.json({});
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.status).toBe(403);
    expect(chunkAttempts).toBe(0);
  });

  it('put rejects an uploadUrl whose host is not on the Microsoft allow-list (Hardening #3)', async () => {
    const total = 6 * 1024 * 1024;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith(':/createUploadSession')) {
        return Response.json({ uploadUrl: 'https://attacker.example.com/exfil' });
      }
      throw new Error('should not have been called');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('not in Microsoft allow-list');
    }
  });

  it('put surfaces api_error when createUploadSession returns no uploadUrl field', async () => {
    const total = 6 * 1024 * 1024;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith(':/createUploadSession')) return Response.json({ surprise: true });
      throw new Error('should not have been called');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('no uploadUrl');
    }
  });

  it('put rejects a malformed uploadUrl from createUploadSession as a network_error', async () => {
    const total = 6 * 1024 * 1024;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith(':/createUploadSession')) return Response.json({ uploadUrl: 'not-a-url' });
      throw new Error('should not have been called');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('invalid uploadUrl');
    }
  });

  it('put cleans up the session and surfaces network_error when a chunk PUT throws', async () => {
    const total = 6 * 1024 * 1024;
    let cleanupCalled = false;
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith(':/createUploadSession')) {
        return Response.json({ uploadUrl: 'https://contoso-my.sharepoint.com/upload-session/throw' });
      }
      if (init?.method === 'DELETE') {
        cleanupCalled = true;
        return new Response(null, { status: 204 });
      }
      throw new Error('disconnected');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('disconnected');
    }
    expect(cleanupCalled).toBe(true);
  });

  it('put simple-path surfaces api_error when Graph rejects with a body containing error.message', async () => {
    const fetchFn: FetchFn = async () => new Response(JSON.stringify({ error: { message: 'name conflict' } }), { status: 409, headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(100));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(409);
      expect(result.error.message).toBe('name conflict');
    }
  });

  it('put simple-path wraps network errors via networkErrorMessage', async () => {
    const fetchFn: FetchFn = async () => {
      throw new Error('boom');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(100));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('boom');
    }
  });

  it('put surfaces auth_failed when the auth manager fails (simple path)', async () => {
    const failingAuth: AuthManager = {
      getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'no token' } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failingAuth);
    const result = await client.put('/me/drive/root:/x', new Uint8Array(100));
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('delete returns ok on a 204 No Content response', async () => {
    const fetchFn: FetchFn = async () => new Response(null, { status: 204 });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.delete('/me/drive/items/i1');
    expect(result.ok).toBe(true);
  });

  it('delete returns api_error on a non-2xx response', async () => {
    const fetchFn: FetchFn = async () => new Response(JSON.stringify({ error: { message: 'cannot delete' } }), { status: 403, headers: { 'content-type': 'application/json' } });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.delete('/me/drive/items/i1');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(403);
      expect(result.error.message).toBe('cannot delete');
    }
  });

  it('delete wraps thrown fetch errors as network_error, labeled DELETE <path>', async () => {
    const fetchFn: FetchFn = async () => {
      throw new Error('reset');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.delete('/me/drive/items/i1');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('reset');
      expect(result.error.message).toContain('DELETE /me/drive/items/i1');
    }
  });

  it('delete surfaces auth_failed when the auth manager fails', async () => {
    const failingAuth: AuthManager = {
      getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'no token' } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failingAuth);
    const result = await client.delete('/me/drive/items/i1');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('chunkedPut surfaces api_error when the final chunk returns 202 instead of 200/201 (truncated stream)', async () => {
    const total = 5 * 1024 * 1024 + 100;
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith(':/createUploadSession')) {
        return Response.json({ uploadUrl: 'https://contoso-my.sharepoint.com/session/abc' });
      }
      if (init?.method === 'DELETE') return new Response(null, { status: 204 });
      return new Response(JSON.stringify({ nextExpectedRanges: ['x-y'] }), { status: 202, headers: { 'content-type': 'application/json' } });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.put('/me/drive/root:/.ask-marcel-temp/x', new Uint8Array(total));
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('chunked upload completed without final response');
    }
  });

  it('getBinaryElevated signs requests with the elevated token (different audience claim than getBinary)', async () => {
    let capturedAuth = '';
    const fetchFn: FetchFn = async (_url, init) => {
      const headers = new Headers(init?.headers);
      capturedAuth = headers.get('authorization') ?? '';
      return new Response(JSON.stringify({ ok: true }), { status: 200, headers: { 'content-type': 'application/json' } });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    await client.getBinaryElevated('/drives/d1/items/i1/versions/2.0/content');
    expect(capturedAuth).toBe('Bearer test-elevated-token');
  });

  it('getBinaryElevated surfaces auth_failed when the elevated auth manager rejects', async () => {
    const failingAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe('test')),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_failed' as const, message: 'elevated capture timed out' } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failingAuth, async () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } }));
    const result = await client.getBinaryElevated('/drives/d1/items/i1/versions/2.0/content');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture timed out');
    }
  });

  it('getBinaryElevated maps auth_cancelled to a friendly auth_failed message', async () => {
    const cancelledAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe('test')),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(cancelledAuth, async () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } }));
    const result = await client.getBinaryElevated('/drives/d1/items/i1/versions/2.0/content');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('Auth cancelled');
    }
  });

  it('getElevated signs JSON GET requests with the elevated token and returns parsed JSON on success', async () => {
    let capturedAuth = '';
    const fetchFn: FetchFn = async (_url, init) => {
      const headers = new Headers(init?.headers);
      capturedAuth = headers.get('authorization') ?? '';
      return new Response(JSON.stringify({ value: [{ id: 'chat-1' }] }), { status: 200, headers: { 'content-type': 'application/json' } });
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getElevated('/me/chats');
    expect(capturedAuth).toBe('Bearer test-elevated-token');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'chat-1' }] });
  });

  it('getElevated surfaces api_error from non-2xx Graph responses', async () => {
    const fetchFn: FetchFn = async () =>
      new Response(JSON.stringify({ error: { code: 'Forbidden', message: 'Insufficient privileges' } }), {
        status: 403,
        headers: { 'content-type': 'application/json' },
      });
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getElevated('/me/chats');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(403);
    }
  });

  it('getElevated surfaces auth_failed when the elevated auth manager rejects', async () => {
    const failingAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe('test')),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_failed' as const, message: 'elevated capture timed out' } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failingAuth, async () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } }));
    const result = await client.getElevated('/me/chats');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture timed out');
    }
  });

  it('getElevated maps auth_cancelled to a friendly auth_failed message', async () => {
    const cancelledAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe('test')),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(cancelledAuth, async () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } }));
    const result = await client.getElevated('/me/chats');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('Auth cancelled');
    }
  });

  it('getElevated maps fetch network failures to network_error', async () => {
    const fetchFn: FetchFn = async () => {
      throw new Error('socket reset');
    };
    const client = createGraphClient(fakeAuth(), fetchFn);
    const result = await client.getElevated('/me/chats');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toContain('socket reset');
    }
  });

  it('getCachedTokenInfo decodes scp/aud/exp claims from the cached token without making a Graph call', async () => {
    const payload = { scp: 'Mail.Read Files.Read User.Read', aud: 'https://graph.microsoft.com', exp: 1893456000 };
    const segment = (s: string): string => Buffer.from(s, 'utf-8').toString('base64').replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
    const jwt = `${segment('{"alg":"none"}')}.${segment(JSON.stringify(payload))}.sig`;
    const tokenAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe(jwt)),
      getElevatedAccessToken: async () => ok(accessTokenUnsafe('not-used')),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(tokenAuth);
    const result = await client.getCachedTokenInfo();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.scopes).toEqual(['Mail.Read', 'Files.Read', 'User.Read']);
      expect(result.value.audience).toBe('https://graph.microsoft.com');
      expect(result.value.expiresAt).toBe(new Date(1893456000 * 1000).toISOString());
    }
  });

  it('getCachedTokenInfo returns empty scopes / undefined audience+expiry when the token has none of those claims', async () => {
    const segment = (s: string): string => Buffer.from(s, 'utf-8').toString('base64').replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
    const jwt = `${segment('{"alg":"none"}')}.${segment('{"sub":"me"}')}.sig`;
    const tokenAuth: AuthManager = {
      getAccessToken: async () => ok(accessTokenUnsafe(jwt)),
      getElevatedAccessToken: async () => ok(accessTokenUnsafe('not-used')),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(tokenAuth);
    const result = await client.getCachedTokenInfo();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.scopes).toEqual([]);
      expect(result.value.audience).toBeUndefined();
      expect(result.value.expiresAt).toBeUndefined();
    }
  });

  it('getCachedTokenInfo returns auth_failed when the auth manager has no token', async () => {
    const cancelledAuth: AuthManager = {
      getAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(cancelledAuth);
    const result = await client.getCachedTokenInfo();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') expect(result.error.message).toBe('Auth cancelled');
  });

  it('getCachedTokenInfo surfaces the auth-manager error message when the failure is not a cancellation', async () => {
    const failedAuth: AuthManager = {
      getAccessToken: async () => ({ ok: false as const, error: { type: 'auth_failed' as const, message: 'token store unreadable' } }),
      getElevatedAccessToken: async () => ({ ok: false as const, error: { type: 'auth_cancelled' as const } }),
      logout: async () => ok(undefined),
      getLastElevatedOutcome: () => null,
    };
    const client = createGraphClient(failedAuth);
    const result = await client.getCachedTokenInfo();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') expect(result.error.message).toBe('token store unreadable');
  });
});
