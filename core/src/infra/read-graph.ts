import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import { tenantId } from '../domain/tenant-id.ts';
import { spoHostToTenantDomain } from '../domain/utilities/spo-tenant.ts';
import type { ReadOnlyPostPath } from '../domain/read-only-post.ts';
import { postIfReadOnly } from '../domain/read-only-post.ts';
import type { ReadGraph } from '../use-cases/ports/graph.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { createGraphRequestCore, wrapNetworkError } from './graph-request.ts';
import { createSubstrateReader } from './graph-substrate.ts';
import { REQUEST_TIMEOUT_MS } from './network-error.ts';

/**
 * Ask Entra which tenant owns a SharePoint host, using its public OIDC
 * discovery document. No credentials: the question is "who owns this host?",
 * and the answer is public.
 *
 * A host outside the `*.sharepoint.com` convention, or a domain Entra does not
 * know, is not an error to retry — it means no partner tenant applies, and the
 * caller should stay on its home token. Both surface as a clear message rather
 * than a crash, because the host->onmicrosoft mapping is a convention and a
 * tenant with a vanity arrangement may not follow it.
 */
const discoverTenantIdWith = async (fetchFn: FetchFn, spoHost: string): Promise<Result<TenantId, GraphError>> => {
  const domain = spoHostToTenantDomain(spoHost);
  if (domain === null) {
    return err({
      type: 'api_error',
      status: 400,
      message: `${spoHost} is not a tenant SharePoint host, so no partner tenant can be resolved from it`,
      code: 'not_a_sharepoint_host',
    });
  }
  const url = `https://login.microsoftonline.com/${domain}/v2.0/.well-known/openid-configuration`;
  try {
    const res = await fetchFn(url, { method: 'GET', headers: { accept: 'application/json' }, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
    if (!res.ok) {
      return err({
        type: 'api_error',
        status: res.status,
        message: `could not resolve a tenant for ${spoHost} (tried ${domain}) — the host may belong to a tenant whose sign-in domain differs from its SharePoint name`,
        code: 'tenant_discovery_failed',
      });
    }
    const issuer = (await res.json())['issuer'];
    // The tenant id is the first path segment of the issuer
    // (`https://login.microsoftonline.com/{tid}/v2.0`). Brand it: it becomes an
    // authority segment on a POST that carries the refresh token.
    const segment = typeof issuer === 'string' ? (new URL(issuer).pathname.split('/')[1] ?? '') : '';
    const branded = tenantId(segment);
    if (!branded.ok) return err({ type: 'api_error', status: 502, message: `tenant discovery for ${spoHost} returned an unusable issuer`, code: 'tenant_discovery_failed' });
    return ok(branded.value);
  } catch (e: unknown) {
    return err(wrapNetworkError(e, 'GET', `tenant discovery for ${spoHost}`, 'json'));
  }
};

const createReadGraph = (tokens: TokenSource, fetchFn: FetchFn = globalThis.fetch): ReadGraph => {
  const core = createGraphRequestCore(tokens, fetchFn);
  return {
    get: (path, extraHeaders) => core.json('basic', 'GET', path, { extraHeaders }),
    getElevated: (path) => core.json('elevated', 'GET', path),
    getGuest: (path, tenant) => core.json({ guest: tenant }, 'GET', path),
    getBinaryGuest: (path, tenant) => core.getBinary({ guest: tenant }, path),
    discoverTenantId: (spoHost) => discoverTenantIdWith(fetchFn, spoHost),
    ...createSubstrateReader(tokens, fetchFn),
    post: (path, body) => postIfReadOnly((readOnly, query) => core.json('basic', 'POST', readOnly, { body: query }), path, body),
    getBinary: (path) => core.getBinary('basic', path),
    getBinaryElevated: (path) => core.getBinary('elevated', path),
    fetchUrl: core.fetchUrl,
  };
};

export { createReadGraph };
export type { ReadGraph, ReadOnlyPostPath };
