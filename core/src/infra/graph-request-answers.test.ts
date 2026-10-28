import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { Result } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { createReadGraph } from './read-graph.ts';

/*
 * How the Graph request core reads a failed answer, through the read graph
 * (core, package split phase 3): the error envelope an agent can act on, a
 * token that cannot be had, and the label of a network failure.
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

const getAnswering = (answer: () => Response, path = '/me/messages'): Promise<Result<unknown, GraphError>> => createReadGraph(tokens, answering(answer)).get(path);

describe('a Graph error, as an agent reads it', () => {
  it('names the code and the message, and keeps the code apart so an agent can branch on it', async () => {
    expect(await getAnswering(() => graphError(404, { code: 'itemNotFound', message: 'The resource could not be found.' }))).toEqual(
      err({ type: 'api_error', status: 404, message: 'itemNotFound: The resource could not be found.', code: 'itemNotFound' })
    );
  });

  it('takes the inner code over the outer one, in the spelling of Graph and of SharePoint', async () => {
    const graph = await getAnswering(() => graphError(410, { code: 'generalException', message: 'Resync.', innererror: { code: 'resyncRequired' } }));
    expect(graph).toEqual(err({ type: 'api_error', status: 410, message: 'resyncRequired: Resync.', code: 'resyncRequired' }));
    const sharePoint = await getAnswering(() => graphError(403, { code: 'accessDenied', message: 'No.', innerError: { code: 'logicalPermissionAccessDenied' } }));
    expect(sharePoint).toEqual(err({ type: 'api_error', status: 403, message: 'logicalPermissionAccessDenied: No.', code: 'logicalPermissionAccessDenied' }));
  });

  it('turns an UnknownError with an empty message into advice to retry once', async () => {
    expect(await getAnswering(() => graphError(500, { code: 'UnknownError', message: '  ' }))).toEqual(
      err({
        type: 'api_error',
        status: 500,
        message:
          'UnknownError: (Graph returned an empty error body — likely a transient backend glitch; retry once. If persistent, capture the failing request URL + body and report.)',
        code: 'UnknownError',
      })
    );
    expect(await getAnswering(() => graphError(500, { code: 'UnknownError' }))).toEqual(
      err(expect.objectContaining({ code: 'UnknownError', message: expect.stringContaining('retry once') }))
    );
  });

  it('keeps an UnknownError that carries a message as it is', async () => {
    expect(await getAnswering(() => graphError(500, { code: 'UnknownError', message: 'Backend busy.' }))).toEqual(
      err({ type: 'api_error', status: 500, message: 'UnknownError: Backend busy.', code: 'UnknownError' })
    );
  });

  it('drops the granted scopes of a missing-scope refusal and names the commands that show them', async () => {
    const refusal = graphError(403, {
      code: 'AccessDenied',
      message: "Missing scope permissions on the request. API requires one of 'Sites.Read.All'. Scopes on the request 'Mail.Read,User.Read'",
    });
    expect(await getAnswering(() => refusal)).toEqual(
      err({
        type: 'api_error',
        status: 403,
        message:
          "AccessDenied: Missing scope permissions on the request. API requires one of 'Sites.Read.All'. Run `{auth-bin} status` to see granted scopes, or `{bin} help-json | jq '.commands[] | select(.name==\"<cmd>\") | .scopesRequired'` to see what a given command requires.",
        code: 'AccessDenied',
      })
    );
  });

  it('gives the message alone, with no code, when Graph sends an empty code', async () => {
    expect(await getAnswering(() => graphError(400, { code: '', message: 'The plan is not valid.' }))).toEqual(
      err({ type: 'api_error', status: 400, message: 'The plan is not valid.' })
    );
  });

  it('falls back to the status text, then to the status and the path, when the body says nothing', async () => {
    expect(await getAnswering(() => new Response(null, { status: 503, statusText: 'Service Unavailable' }))).toEqual(
      err({ type: 'api_error', status: 503, message: 'Service Unavailable' })
    );
    expect(await getAnswering(() => new Response('not json', { status: 404 }), '/me/insights/trending')).toEqual(
      err({
        type: 'api_error',
        status: 404,
        message: 'HTTP 404 with no error body (path: /v1.0/me/insights/trending; the endpoint may have moved — see the command\'s "best-effort" note in --help)',
      })
    );
  });

  it('marks a malformed id on a mail folder path, so the hint can name the well-known folders', async () => {
    const malformed = (): Response => graphError(400, { code: 'ErrorInvalidIdMalformed', message: 'Id is malformed.' });
    expect(await getAnswering(malformed, '/me/mailFolders/inbx/messages')).toEqual(err(expect.objectContaining({ code: 'ErrorInvalidIdMalformed_mailFolders' })));
    expect(await getAnswering(malformed, "/me/messages?$filter=parentFolderId eq 'mailFolders%2Finbx'")).toEqual(
      err(expect.objectContaining({ code: 'ErrorInvalidIdMalformed_mailFolders' }))
    );
    expect(await getAnswering(() => graphError(400, { code: 'InvalidIdMalformed', message: 'Id is malformed.' }), '/me/mailFolders/x')).toEqual(
      err(expect.objectContaining({ code: 'InvalidIdMalformed_mailFolders' }))
    );
    expect(await getAnswering(malformed, '/me/messages/abc')).toEqual(err(expect.objectContaining({ code: 'ErrorInvalidIdMalformed' })));
    expect(await getAnswering(() => graphError(400, { code: 'BadRequest', message: 'No.' }), '/me/mailFolders/x')).toEqual(err(expect.objectContaining({ code: 'BadRequest' })));
  });

  it('gives the seconds Graph asks to wait before a retry, when it names a whole number', async () => {
    const createThrottledAnswer = (retryAfter: string) => (): Response => graphError(429, { code: 'TooManyRequests', message: 'Slow down.' }, { 'retry-after': retryAfter });
    expect(await getAnswering(createThrottledAnswer(' 7 '))).toEqual(
      err({ type: 'api_error', status: 429, message: 'TooManyRequests: Slow down.', code: 'TooManyRequests', retryAfterSeconds: 7 })
    );
    expect(await getAnswering(createThrottledAnswer('0'))).toEqual(err(expect.objectContaining({ retryAfterSeconds: 0 })));
    expect(await getAnswering(createThrottledAnswer('Wed, 21 Oct 2026 07:28:00 GMT'))).toEqual(
      err({ type: 'api_error', status: 429, message: 'TooManyRequests: Slow down.', code: 'TooManyRequests' })
    );
    expect(await getAnswering(createThrottledAnswer('1.5'))).toEqual(err(expect.not.objectContaining({ retryAfterSeconds: expect.anything() })));
    const unknown = await getAnswering(() => graphError(503, { code: 'UnknownError' }, { 'retry-after': '3' }));
    expect(unknown).toEqual(err(expect.objectContaining({ code: 'UnknownError', retryAfterSeconds: 3 })));
    const bare = await getAnswering(() => new Response(null, { status: 503, headers: { 'retry-after': '4' } }), '/me/events');
    expect(bare).toEqual(err(expect.objectContaining({ retryAfterSeconds: 4 })));
  });
});

describe('a request no token could sign', () => {
  it('reports a cancelled sign-in as an auth failure with no code', async () => {
    const cancelled: TokenSource = { ...tokens, graphToken: async () => err({ type: 'auth_cancelled' }) };
    expect(
      await createReadGraph(
        cancelled,
        answering(() => Response.json({}))
      ).get('/me')
    ).toEqual(err({ type: 'auth_failed', message: 'Auth cancelled' }));
  });

  it('keeps the code of the auth layer, and gives none when it has none', async () => {
    const coded: TokenSource = { ...tokens, guestToken: async () => err({ type: 'auth_failed', message: 'No guest token.', code: 'secondary_token_unavailable' }) };
    expect(
      await createReadGraph(
        coded,
        answering(() => Response.json({}))
      ).getGuest('/drives/d', PARTNER)
    ).toEqual(err({ type: 'auth_failed', message: 'No guest token.', code: 'secondary_token_unavailable' }));
    const bare: TokenSource = { ...tokens, graphToken: async () => err({ type: 'auth_failed', message: 'No token.' }) };
    expect(
      await createReadGraph(
        bare,
        answering(() => Response.json({}))
      ).getElevated('/me')
    ).toEqual(err({ type: 'auth_failed', message: 'No token.' }));
  });
});

describe('a network failure', () => {
  const message = (base: string, label: string): string => `${base} (${label}) — transient; retry once before treating as permanent`;

  it('names the call and the tier it was signed for', async () => {
    expect(await createReadGraph(tokens, failing(new Error('fetch failed'))).get('/me')).toEqual(err({ type: 'network_error', message: message('fetch failed', 'GET /me') }));
    expect(await createReadGraph(tokens, failing(new Error('fetch failed'))).getElevated('/me')).toEqual(
      err({ type: 'network_error', message: message('fetch failed', 'GET /me (elevated)') })
    );
    expect(await createReadGraph(tokens, failing(new Error('fetch failed'))).getGuest('/drives/d', PARTNER)).toEqual(
      err({ type: 'network_error', message: message('fetch failed', `GET /drives/d (guest ${PARTNER})`) })
    );
  });

  it('says what stopped the request: an abort, a thrown text, or nothing known', async () => {
    expect(await createReadGraph(tokens, failing(new DOMException('aborted', 'AbortError'))).get('/me')).toEqual(
      err({ type: 'network_error', message: message('request aborted', 'GET /me') })
    );
    expect(await createReadGraph(tokens, failing('socket closed')).get('/me')).toEqual(err({ type: 'network_error', message: message('socket closed', 'GET /me') }));
    expect(await createReadGraph(tokens, failing({ reason: 'odd' })).get('/me')).toEqual(err({ type: 'network_error', message: message('network request failed', 'GET /me') }));
  });

  it('names a file read and the file-transfer deadline', async () => {
    expect(await createReadGraph(tokens, failing(new DOMException('timed out', 'TimeoutError'))).getBinary('/me/drive/items/i/content')).toEqual(
      err({ type: 'network_error', message: message('request timed out after 5min', 'GET /me/drive/items/i/content (binary)') })
    );
  });
});
