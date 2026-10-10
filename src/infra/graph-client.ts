import type { AccessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AuthManager } from '../infra/auth.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import { createAuthManagerTokenSource } from './auth-token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { apiErrorFrom, asAuthFailure, createGraphRequestCore, isAllowedFetchUrlHost, wrapNetworkError } from './graph-request.ts';
import type { ReadGraph } from './read-graph.ts';
import { createReadGraph } from './read-graph.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';

// The client: every read of the read graph, and the writes the write
// commands use. Its POST reaches any path.
type GraphClient = ReadGraph & {
  post: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  patch: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
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
    ...createReadGraph(tokens, fetchFn),
    post: (path, body) => request('POST', path, body),
    patch: (path, body) => request('PATCH', path, body),
    put,
    delete: deleteResource,
  };
};

// The default: the in-process auth ladder, or a library caller's own manager.
const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => createTokenSourceGraphClient(createAuthManagerTokenSource(auth), fetchFn);

export { createGraphClient, createTokenSourceGraphClient };
export type { FetchFn, GraphClient, GraphError };
