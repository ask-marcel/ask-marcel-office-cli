import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { createReadGraph } from './read-graph.ts';

/*
 * The read graph's reads beyond Graph v1.0 (core, package split phase 3): the
 * tenant that owns a SharePoint host, the Teams chat substrate on chatsvcagg
 * and IC3, and a pasted chat image from the Teams media service.
 */

type Sent = { readonly url: string; readonly authorization: string | null };

const recording = (sent: Sent[], answers: ReadonlyArray<() => Response>): FetchFn => {
  const queue = [...answers];
  const fetchFn: FetchFn = async (url, init) => {
    sent.push({ url, authorization: new Headers(init?.headers).get('authorization') });
    const answer = queue.shift();
    if (answer === undefined) throw new Error('no more answers');
    return answer();
  };
  return fetchFn;
};

const failing = (thrown: unknown): FetchFn => {
  const fetchFn: FetchFn = async () => {
    throw thrown;
  };
  return fetchFn;
};

// A substrate source that mints `fresh-N` tokens, each past the one it is told was refused.
const substrateTokens = (asked: Array<{ readonly tier: string; readonly rejected?: AccessToken }>): TokenSource => ({
  graphToken: async () => ok(accessTokenUnsafe('basic-token')),
  guestToken: async () => ok(accessTokenUnsafe('guest-token')),
  substrateToken: async (tier, options) => {
    asked.push({ tier, ...(options?.rejected === undefined ? {} : { rejected: options.rejected }) });
    return ok(accessTokenUnsafe(`fresh-${asked.length}`));
  },
  substrateRegion: async () => ok('amer' as TeamsRegion),
});

const ISSUER = 'https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47/v2.0';

describe('the tenant that owns a SharePoint host', () => {
  const discovery = (domain: string): string => `https://login.microsoftonline.com/${domain}/v2.0/.well-known/openid-configuration`;

  it('asks Entra about the onmicrosoft domain behind a site, a personal OneDrive or an admin host, with no bearer', async () => {
    for (const [host, domain] of [
      ['contoso.sharepoint.com', 'contoso.onmicrosoft.com'],
      ['Contoso-My.SharePoint.com', 'contoso.onmicrosoft.com'],
      ['fabrikam-admin.sharepoint.com', 'fabrikam.onmicrosoft.com'],
      ['north-wind-my.sharepoint.com', 'north-wind.onmicrosoft.com'],
      ['a1.sharepoint.com', 'a1.onmicrosoft.com'],
    ] as const) {
      const sent: Sent[] = [];
      const graph = createReadGraph(substrateTokens([]), recording(sent, [() => Response.json({ issuer: ISSUER })]));
      expect(await graph.discoverTenantId(host)).toEqual(ok(tenantIdUnsafe('72f988bf-86f1-41af-91ab-2d7cd011db47')));
      expect(sent).toEqual([{ url: discovery(domain), authorization: null }]);
    }
  });

  it('refuses a host that is not a tenant SharePoint host, before any request', async () => {
    for (const host of [
      '1drv.ms',
      'contoso.sharepoint.com.evil.example',
      'a.b.sharepoint.com',
      '-x.sharepoint.com',
      'sharepoint.com',
      'contoso.sharepoint.org',
      'xcontoso.sharepoint.comx',
    ]) {
      const sent: Sent[] = [];
      expect(await createReadGraph(substrateTokens([]), recording(sent, [])).discoverTenantId(host)).toEqual(
        err({ type: 'api_error', status: 400, message: `${host} is not a tenant SharePoint host, so no partner tenant can be resolved from it`, code: 'not_a_sharepoint_host' })
      );
      expect(sent).toEqual([]);
    }
  });

  it('says the tenant could not be resolved when Entra does not know the domain', async () => {
    const graph = createReadGraph(substrateTokens([]), recording([], [() => Response.json({ error: 'invalid_tenant' }, { status: 400 })]));
    expect(await graph.discoverTenantId('vanity.sharepoint.com')).toEqual(
      err({
        type: 'api_error',
        status: 400,
        message:
          'could not resolve a tenant for vanity.sharepoint.com (tried vanity.onmicrosoft.com) — the host may belong to a tenant whose sign-in domain differs from its SharePoint name',
        code: 'tenant_discovery_failed',
      })
    );
  });

  it('refuses an issuer that names no tenant id, or no issuer at all', async () => {
    const unusable: Result<never, GraphError> = err({
      type: 'api_error',
      status: 502,
      message: 'tenant discovery for contoso.sharepoint.com returned an unusable issuer',
      code: 'tenant_discovery_failed',
    });
    for (const body of [{ issuer: 'https://login.microsoftonline.com/common/v2.0' }, { issuer: 42 }, {}]) {
      expect(await createReadGraph(substrateTokens([]), recording([], [() => Response.json(body)])).discoverTenantId('contoso.sharepoint.com')).toEqual(unusable);
    }
  });

  it('names the host when the network fails', async () => {
    expect(await createReadGraph(substrateTokens([]), failing(new Error('fetch failed'))).discoverTenantId('contoso.sharepoint.com')).toEqual(
      err({ type: 'network_error', message: 'fetch failed (GET tenant discovery for contoso.sharepoint.com) — transient; retry once before treating as permanent' })
    );
  });
});

describe('a Teams chat read', () => {
  it('reads chatsvcagg and IC3 under the region of the token, each with its own bearer', async () => {
    const asked: Array<{ readonly tier: string }> = [];
    const sent: Sent[] = [];
    const graph = createReadGraph(substrateTokens(asked), recording(sent, [() => Response.json({ chats: [] }), () => Response.json({ messages: [] })]));
    expect(await graph.teamsChat('/chats')).toEqual(ok({ chats: [] }));
    expect(await graph.teamsChatIc3('/users/ME/conversations')).toEqual(ok({ messages: [] }));
    expect(sent).toEqual([
      { url: 'https://teams.microsoft.com/api/csa/amer/chats', authorization: 'Bearer fresh-1' },
      { url: 'https://teams.microsoft.com/api/chatsvc/amer/users/ME/conversations', authorization: 'Bearer fresh-2' },
    ]);
    expect(asked).toEqual([{ tier: 'chatsvcagg' }, { tier: 'ic3' }]);
  });

  it('replays a 401 once with a token past the refused one, and surfaces a second 401', async () => {
    const asked: Array<{ readonly tier: string; readonly rejected?: AccessToken }> = [];
    const sent: Sent[] = [];
    const refused = (): Response => new Response(null, { status: 401 });
    const graph = createReadGraph(substrateTokens(asked), recording(sent, [refused, refused]));
    expect(await graph.teamsChat('/chats')).toEqual(err(expect.objectContaining({ type: 'api_error', status: 401, code: 'substrateHttp401_chatsvcagg' })));
    expect(asked).toEqual([{ tier: 'chatsvcagg' }, { tier: 'chatsvcagg', rejected: accessTokenUnsafe('fresh-1') }]);
    expect(sent.map((s) => s.authorization)).toEqual(['Bearer fresh-1', 'Bearer fresh-2']);
  });

  it('answers with the replay when the fresh token is taken', async () => {
    const graph = createReadGraph(substrateTokens([]), recording([], [() => new Response(null, { status: 401 }), () => Response.json({ ok: 1 })]));
    expect(await graph.teamsChatIc3('/x')).toEqual(ok({ ok: 1 }));
  });

  it('tags a substrate error with its status and substrate, so the hint can say the substrate may have moved', async () => {
    const graph = createReadGraph(substrateTokens([]), recording([], [() => new Response(null, { status: 404, statusText: 'Not Found' })]));
    expect(await graph.teamsChatIc3('/x')).toEqual(err({ type: 'api_error', status: 404, message: 'Not Found', code: 'substrateHttp404_ic3' }));
  });

  it('reports a token or region that cannot be had as an auth failure, and sends nothing', async () => {
    const sent: Sent[] = [];
    const noRegion: TokenSource = { ...substrateTokens([]), substrateRegion: async () => err({ type: 'auth_failed', message: 'No region.', code: 'secondary_token_unavailable' }) };
    expect(await createReadGraph(noRegion, recording(sent, [])).teamsChat('/x')).toEqual(err({ type: 'auth_failed', message: 'No region.', code: 'secondary_token_unavailable' }));
    const noToken: TokenSource = { ...substrateTokens([]), substrateToken: async () => err({ type: 'auth_cancelled' }) };
    expect(await createReadGraph(noToken, recording(sent, [])).teamsChatIc3('/x')).toEqual(err({ type: 'auth_failed', message: 'Auth cancelled' }));
    expect(sent).toEqual([]);
  });

  it('reports a replay whose token cannot be had as an auth failure', async () => {
    let calls = 0;
    const once: TokenSource = {
      ...substrateTokens([]),
      substrateToken: async () => {
        calls += 1;
        return calls === 1 ? ok(accessTokenUnsafe('stale')) : err({ type: 'auth_failed', message: 'Refresh failed.', code: 'secondary_token_unavailable' });
      },
    };
    const graph = createReadGraph(once, recording([], [() => new Response(null, { status: 401 })]));
    expect(await graph.teamsChat('/x')).toEqual(err({ type: 'auth_failed', message: 'Refresh failed.', code: 'secondary_token_unavailable' }));
  });

  it('keeps a failure that is not a Graph answer as it is', async () => {
    expect(await createReadGraph(substrateTokens([]), failing(new Error('fetch failed'))).teamsChat('/chats')).toEqual(
      err({ type: 'network_error', message: 'fetch failed (GET /chats (chatsvcagg)) — transient; retry once before treating as permanent' })
    );
  });
});

describe('a pasted Teams chat image', () => {
  const MEDIA = 'https://eu-api.asm.skype.com/v1/objects/0-weu-d1-abc/views/imgo';

  it('is read with the IC3 bearer as bytes, from either Teams media host', async () => {
    const sent: Sent[] = [];
    const image = (): Response => new Response(new Uint8Array([137, 80]), { headers: { 'content-type': 'image/png' } });
    const graph = createReadGraph(substrateTokens([]), recording(sent, [image, image]));
    expect(await graph.teamsChatMedia(MEDIA)).toEqual(ok({ contentType: 'image/png', size: 2, base64: 'iVA=' }));
    expect(await graph.teamsChatMedia('https://us-prod.asyncgw.teams.microsoft.com/v1/objects/x/views/imgo')).toEqual(ok({ contentType: 'image/png', size: 2, base64: 'iVA=' }));
    expect(sent.map((s) => s.authorization)).toEqual(['Bearer fresh-1', 'Bearer fresh-2']);
  });

  it('names an image with no content type as plain bytes', async () => {
    const untyped = (): Response => {
      const response = new Response(new Uint8Array([1]));
      response.headers.delete('content-type');
      return response;
    };
    expect(await createReadGraph(substrateTokens([]), recording([], [untyped])).teamsChatMedia(MEDIA)).toEqual(
      ok({ contentType: 'application/octet-stream', size: 1, base64: 'AQ==' })
    );
  });

  it('never sends the bearer to another host, to plain http, or to a text that is not a URL', async () => {
    const sent: Sent[] = [];
    const graph = createReadGraph(substrateTokens([]), recording(sent, []));
    for (const url of [MEDIA.replace('https', 'http'), 'https://evil.example/asm.skype.com', 'https://asm.skype.com.evil.example/x', 'https://a.b.asm.skype.com/x', 'not a url']) {
      expect(await graph.teamsChatMedia(url)).toEqual(
        err({ type: 'validation_error', message: `not a Teams media URL (https on *.asm.skype.com or *.asyncgw.teams.microsoft.com): ${url}` })
      );
    }
    expect(sent).toEqual([]);
  });

  it('quotes at most 120 characters of a refused URL', async () => {
    const long = `https://evil.example/${'x'.repeat(200)}`;
    expect(await createReadGraph(substrateTokens([]), recording([], [])).teamsChatMedia(long)).toEqual(
      err({ type: 'validation_error', message: `not a Teams media URL (https on *.asm.skype.com or *.asyncgw.teams.microsoft.com): ${long.slice(0, 120)}` })
    );
  });

  it('reports a refused image as the answer, and a failed read with the host and the file-transfer deadline', async () => {
    expect(await createReadGraph(substrateTokens([]), recording([], [() => new Response(null, { status: 403, statusText: 'Forbidden' })])).teamsChatMedia(MEDIA)).toEqual(
      err({ type: 'api_error', status: 403, message: 'Forbidden' })
    );
    expect(await createReadGraph(substrateTokens([]), failing(new DOMException('timed out', 'TimeoutError'))).teamsChatMedia(MEDIA)).toEqual(
      err({ type: 'network_error', message: 'request timed out after 5min (GET eu-api.asm.skype.com (teams media)) — transient; retry once before treating as permanent' })
    );
  });

  it('reports a token that cannot be had as an auth failure', async () => {
    const noToken: TokenSource = { ...substrateTokens([]), substrateToken: async () => err({ type: 'auth_failed', message: 'No ic3 token.' }) };
    expect(await createReadGraph(noToken, recording([], [])).teamsChatMedia(MEDIA)).toEqual(err({ type: 'auth_failed', message: 'No ic3 token.' }));
  });
});
