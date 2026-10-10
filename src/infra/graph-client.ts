import type { AccessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AuthManager } from '../infra/auth.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import { tenantId } from '../domain/tenant-id.ts';
import { spoHostToTenantDomain } from '../domain/utilities/spo-tenant.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import { createAuthManagerTokenSource } from './auth-token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { apiErrorFrom, asAuthFailure, createGraphRequestCore, isAllowedFetchUrlHost, wrapNetworkError } from './graph-request.ts';
import { createSubstrateReader } from './graph-substrate.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';

type GraphClient = {
  /**
   * `extraHeaders` lets a caller add request headers Graph requires on
   * specific endpoints — currently the only documented use is
   * `Prefer: odata.maxpagesize=N` on the calendar/mail delta endpoints,
   * which reject `$top` as a query parameter. Auth + content-type are
   * always set internally.
   */
  get: (path: string, extraHeaders?: Record<string, string>) => Promise<Result<unknown, GraphError>>;
  /**
   * Same JSON-GET shape as `get`, but signs the request with the
   * elevated Graph token (M365ChatClient). Used by commands the Teams
   * web client token cannot reach — currently `list-chats` and
   * `get-chat`, which need `Chat.ReadBasic` (only present on the
   * elevated token).
   */
  getElevated: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * JSON-GET signed with a PARTNER tenant's guest token instead of the home
   * token. Required for any path that touches a tenant the user is only a guest
   * in: home-tenant Graph cannot mint a SharePoint token for a foreign tenant,
   * so those calls die at `401 invalidAudienceUri` no matter which home tier
   * signs them.
   *
   * Get the `tenantId` from `resolve-drive-share-link` (it discovers it from the
   * sharing URL) or from the caller's `--tenant-id`.
   */
  getGuest: (path: string, tenantId: TenantId) => Promise<Result<unknown, GraphError>>;
  /**
   * Binary twin of `getGuest`: follows the Graph 302 to the partner tenant's CDN
   * and returns the bytes. The `fetchUrl` allow-list already admits any
   * `*.sharepoint.com` / `*.svc.ms` host, so a partner tenant's download URL
   * needs no special casing.
   */
  getBinaryGuest: (path: string, tenantId: TenantId) => Promise<Result<unknown, GraphError>>;
  /**
   * Resolves a SharePoint host to the Entra tenant that owns it, via the tenant's
   * public OIDC discovery document. Unauthenticated: it asks "who owns this
   * host?", not "what may I read?".
   *
   * This is what makes a bare sharing URL enough to cross tenants — the URL
   * carries the host, the host names the tenant, and the tenant is the one thing
   * `driveId` + `itemId` do not tell you.
   */
  discoverTenantId: (spoHost: string) => Promise<Result<TenantId, GraphError>>;
  /**
   * JSON-GET against the Teams chat substrate (post-2026-05:
   * `teams.microsoft.com/api/csa/<region>/api/v{N}/...` — see
   * `gotcha_chatsvcagg_substrate_moved` in memory for the migration
   * away from `chatsvcagg.teams.microsoft.com`). Signs the request
   * with the chatsvcagg-audience bearer captured at login (same Teams
   * web client identity as `get`, different audience), and injects the
   * cached substrate region between the host and `path`. Used by
   * commands that need to read chat message BODIES, which the basic
   * Graph token cannot reach (`Chat.Read*` scopes are missing).
   *
   * `path` MUST start with `/api/v{N}/...` — the host + `/api/csa/<region>`
   * prefix are added by this client.
   */
  teamsChat: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * JSON-GET against the Teams IC3 chat-message substrate at
   * `teams.microsoft.com/api/chatsvc/<region>/v1/...`. Same host as
   * `teamsChat` but a DIFFERENT path prefix AND a different bearer
   * audience (`https://ic3.teams.office.com` instead of
   * `https://chatsvcagg.teams.microsoft.com`). The path supports
   * `syncState` + `startTime` pagination — unlocking arbitrary-depth
   * chat-history reads beyond the chatsvcagg 200-message cap (see
   * `gotcha_chatsvcagg_substrate_moved` in memory). Used by
   * `list-teams-chat-history`.
   *
   * `path` MUST start with `/v1/...` (e.g. `/v1/users/ME/conversations/{id}/messages?startTime=...`)
   * — the host + `/api/chatsvc/<region>` prefix are added here.
   */
  teamsChatIc3: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * A pasted chat image from Teams' media service (`*.asm.skype.com`,
   * `*.asyncgw.teams.microsoft.com`), read with the IC3 bearer, the only token it
   * accepts (probed 2026-09-27: 200 with it, 401 with every other). The URL comes
   * out of message content, so any other host, or plain http, is refused before a
   * request is made. Answers `{ contentType, size, base64 }`.
   */
  teamsChatMedia: (url: string) => Promise<Result<unknown, GraphError>>;
  post: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  patch: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  getBinary: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Same shape as `getBinary` but signs the request with an "elevated"
   * Graph token (issued for an app on Microsoft's ODSP
   * `logicalPermissions` allow-list — e.g., M365ChatClient). Used by
   * the historical-version commands which the Teams web client token
   * cannot fetch (403 logicalPermissionAccessDenied).
   */
  getBinaryElevated: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Auth-less fetch of an arbitrary URL whose host MUST be on the
   * Microsoft allow-list. Used to follow `@microsoft.graph.downloadUrl`
   * 302 redirects (CDN-signed URLs) that the format-conversion
   * commands sometimes get back from Graph instead of inline bytes.
   */
  fetchUrl: (url: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Upload bytes to a drive item. `basePath` is the bare driveItem
   * path (e.g. `/me/drive/root:/.ask-marcel-temp/abc.rtf`) — `put()`
   * appends `:/content` for the simple ≤4 MiB sync path or
   * `:/createUploadSession` for the chunked-session path internally
   * based on `body.byteLength`. No upper file-size limit beyond the
   * user's OneDrive quota.
   */
  put: (basePath: string, body: Uint8Array, contentType?: string) => Promise<Result<unknown, GraphError>>;
  delete: (path: string) => Promise<Result<unknown, GraphError>>;
};

// Two-tier timeout constants live in src/infra/network-error.ts (shared
// with the TeamsClient adapter). The chunk constants are GraphClient-
// specific so they stay here.
const SIMPLE_PUT_THRESHOLD = 4 * 1024 * 1024; // 4 MiB
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MiB — Graph requires multiples of 320 KiB; 5 MiB is 16 × 320 KiB

// Every bearer comes from the token source, which knows the tiers; this client
// only signs requests with them.
const createTokenSourceGraphClient = (tokens: TokenSource, fetchFn: FetchFn = globalThis.fetch): GraphClient => {
  const bearer = async (token: Promise<Result<AccessToken, TokenError>>): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await token;
    if (!tokenResult.ok) return err(asAuthFailure(tokenResult.error));
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
  };
  const authHeaders = (): Promise<Result<{ Authorization: string }, GraphError>> => bearer(tokens.graphToken('basic'));

  const core = createGraphRequestCore(tokens, fetchFn);
  const request = (method: 'GET' | 'POST' | 'PATCH', path: string, body?: unknown, extraHeaders?: Record<string, string>): Promise<Result<unknown, GraphError>> =>
    core.json('basic', method, path, { body, extraHeaders });

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
  const discoverTenantId = async (spoHost: string): Promise<Result<TenantId, GraphError>> => {
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

  const simplePut = async (path: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'PUT',
        headers: { ...headers.value, 'content-type': contentType ?? 'application/octet-stream' },
        body: body as unknown as BodyInit,
        signal: AbortSignal.timeout(timeoutMsFor('binary')),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'PUT', path, 'binary'));
    }
  };

  const chunkedPut = async (basePath: string, body: Uint8Array): Promise<Result<unknown, GraphError>> => {
    // 1. Create upload session via the authenticated request() helper.
    const session = await request('POST', `${basePath}:/createUploadSession`, {
      item: { '@microsoft.graph.conflictBehavior': 'replace' },
    });
    if (!session.ok) return session;
    const uploadUrl = (session.value as { uploadUrl?: string }).uploadUrl;
    if (typeof uploadUrl !== 'string') {
      return err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
    }

    // Hardening #3: validate the host before any chunk PUT.
    let host: string;
    try {
      host = new URL(uploadUrl).host;
    } catch {
      return err({ type: 'network_error', message: 'createUploadSession returned an invalid uploadUrl' });
    }
    if (!isAllowedFetchUrlHost(host)) {
      return err({ type: 'network_error', message: `uploadUrl host ${host} not in Microsoft allow-list` });
    }

    // 2. PUT chunks to the pre-signed upload URL — no auth header.
    const total = body.byteLength;
    for (let start = 0; start < total; start += CHUNK_SIZE) {
      const end = Math.min(start + CHUNK_SIZE, total) - 1;
      const chunk = body.slice(start, end + 1);
      try {
        const res = await fetchFn(uploadUrl, {
          method: 'PUT',
          headers: { 'Content-Range': `bytes ${start}-${end}/${total}` },
          body: chunk as unknown as BodyInit,
          signal: AbortSignal.timeout(timeoutMsFor('binary')),
        });
        if (!res.ok) {
          // Best-effort session cancellation; ignore failure. DELETE keeps
          // the short-tier budget — it's a Graph-side cleanup that should
          // return promptly.
          try {
            await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
          } catch {
            /* ignore */
          }
          return err({ type: 'api_error', status: res.status, message: `chunk PUT failed at byte ${start}` });
        }
        if (res.status === 200 || res.status === 201) {
          return ok(await res.json());
        }
        // 202 Accepted — continue uploading.
      } catch (e: unknown) {
        try {
          await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
        } catch {
          /* ignore */
        }
        return err(wrapNetworkError(e, 'PUT', `chunk @ byte ${start}`, 'binary'));
      }
    }
    return err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' });
  };

  const put = async (basePath: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    if (body.byteLength <= SIMPLE_PUT_THRESHOLD) {
      return simplePut(`${basePath}:/content`, body, contentType);
    }
    return chunkedPut(basePath, body);
  };

  const deleteResource = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'DELETE',
        headers: headers.value,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(undefined);
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'DELETE', path, 'json'));
    }
  };

  return {
    get: (path, extraHeaders) => request('GET', path, undefined, extraHeaders),
    getElevated: (path) => core.json('elevated', 'GET', path),
    getGuest: (path, tenant) => core.json({ guest: tenant }, 'GET', path),
    discoverTenantId,
    ...createSubstrateReader(tokens, fetchFn),
    post: (path, body) => request('POST', path, body),
    patch: (path, body) => request('PATCH', path, body),
    getBinary: (path) => core.getBinary('basic', path),
    getBinaryElevated: (path) => core.getBinary('elevated', path),
    getBinaryGuest: (path, tenant) => core.getBinary({ guest: tenant }, path),
    fetchUrl: core.fetchUrl,
    put,
    delete: deleteResource,
  };
};

// The default: the in-process auth ladder, or a library caller's own manager.
const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => createTokenSourceGraphClient(createAuthManagerTokenSource(auth), fetchFn);

export { createGraphClient, createTokenSourceGraphClient };
export type { FetchFn, GraphClient, GraphError };
