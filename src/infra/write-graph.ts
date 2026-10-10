import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError, GraphRequestCore } from './graph-request.ts';
import { apiErrorFrom, createGraphRequestCore, GRAPH_ROOT, isAllowedFetchUrlHost, wrapNetworkError } from './graph-request.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';

/*
 * The graph a write command gets (package split, D9): the basic tier only,
 * which is all the write commands sign with. It reads as they need (JSON,
 * bytes, the CDN follow) and adds post, patch, put and delete. Read code never
 * imports this module (eslint.config.js), so a read-only build carries none of
 * it (scripts/check-read-bundle.ts).
 */

type WriteGraph = {
  get: (path: string, extraHeaders?: Record<string, string>) => Promise<Result<unknown, GraphError>>;
  getBinary: (path: string) => Promise<Result<unknown, GraphError>>;
  fetchUrl: (url: string) => Promise<Result<unknown, GraphError>>;
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

const SIMPLE_PUT_THRESHOLD = 4 * 1024 * 1024; // 4 MiB
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MiB — Graph requires multiples of 320 KiB; 5 MiB is 16 × 320 KiB

// A write Graph accepted with no body (202 Accepted, 204 No Content) is done.
// Read as JSON, the empty body would report a network error, and a caller could
// send a write that was not idempotent a second time.
const readWriteAnswer = async (res: Response): Promise<unknown> => {
  const text = await res.text();
  return text === '' ? undefined : JSON.parse(text);
};

// Best-effort session cancellation; ignore failure. DELETE keeps the
// short-tier budget — it's a Graph-side cleanup that should return promptly.
const cancelUploadSession = async (fetchFn: FetchFn, uploadUrl: string): Promise<void> => {
  try {
    await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
  } catch {
    /* ignore */
  }
};

const uploadUrlHost = (uploadUrl: string): string | undefined => {
  try {
    return new URL(uploadUrl).host;
  } catch {
    return undefined;
  }
};

const simplePut = async (core: GraphRequestCore, path: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
  const url = `${GRAPH_ROOT}${path}`;
  try {
    const sent = await core.send('basic', url, {
      method: 'PUT',
      headers: { 'content-type': contentType ?? 'application/octet-stream' },
      body: body as unknown as BodyInit,
      timeoutMs: timeoutMsFor('binary'),
    });
    if (!sent.ok) return sent;
    if (!sent.value.ok) return err(await apiErrorFrom(sent.value, url));
    return ok(await readWriteAnswer(sent.value));
  } catch (e: unknown) {
    return err(wrapNetworkError(e, 'PUT', path, 'binary'));
  }
};

// 2. PUT chunks to the pre-signed upload URL — no auth header.
const putChunks = async (fetchFn: FetchFn, uploadUrl: string, body: Uint8Array): Promise<Result<unknown, GraphError>> => {
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
        await cancelUploadSession(fetchFn, uploadUrl);
        return err({ type: 'api_error', status: res.status, message: `chunk PUT failed at byte ${start}` });
      }
      if (res.status === 200 || res.status === 201) {
        return ok(await res.json());
      }
      // 202 Accepted — continue uploading.
    } catch (e: unknown) {
      await cancelUploadSession(fetchFn, uploadUrl);
      return err(wrapNetworkError(e, 'PUT', `chunk @ byte ${start}`, 'binary'));
    }
  }
  return err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' });
};

const chunkedPut = async (core: GraphRequestCore, fetchFn: FetchFn, basePath: string, body: Uint8Array): Promise<Result<unknown, GraphError>> => {
  // 1. Create upload session via the authenticated request helper.
  const session = await core.json('basic', 'POST', `${basePath}:/createUploadSession`, {
    body: { item: { '@microsoft.graph.conflictBehavior': 'replace' } },
    readBody: readWriteAnswer,
  });
  if (!session.ok) return session;
  const uploadUrl = (session.value as { uploadUrl?: string } | undefined)?.uploadUrl;
  if (typeof uploadUrl !== 'string') {
    return err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
  }

  // Hardening #3: validate the host before any chunk PUT.
  const host = uploadUrlHost(uploadUrl);
  if (host === undefined) return err({ type: 'network_error', message: 'createUploadSession returned an invalid uploadUrl' });
  if (!isAllowedFetchUrlHost(host)) {
    return err({ type: 'network_error', message: `uploadUrl host ${host} not in Microsoft allow-list` });
  }
  return putChunks(fetchFn, uploadUrl, body);
};

const deleteWith = async (core: GraphRequestCore, path: string): Promise<Result<unknown, GraphError>> => {
  const url = `${GRAPH_ROOT}${path}`;
  try {
    const sent = await core.send('basic', url, { method: 'DELETE', timeoutMs: REQUEST_TIMEOUT_MS });
    if (!sent.ok) return sent;
    if (!sent.value.ok) return err(await apiErrorFrom(sent.value, url));
    return ok(undefined);
  } catch (e: unknown) {
    return err(wrapNetworkError(e, 'DELETE', path, 'json'));
  }
};

const createWriteGraph = (tokens: TokenSource, fetchFn: FetchFn = globalThis.fetch): WriteGraph => {
  const core = createGraphRequestCore(tokens, fetchFn);
  return {
    get: (path, extraHeaders) => core.json('basic', 'GET', path, { extraHeaders }),
    getBinary: (path) => core.getBinary('basic', path),
    fetchUrl: core.fetchUrl,
    post: (path, body) => core.json('basic', 'POST', path, { body, readBody: readWriteAnswer }),
    patch: (path, body) => core.json('basic', 'PATCH', path, { body, readBody: readWriteAnswer }),
    put: (basePath, body, contentType) =>
      body.byteLength <= SIMPLE_PUT_THRESHOLD ? simplePut(core, `${basePath}:/content`, body, contentType) : chunkedPut(core, fetchFn, basePath, body),
    delete: (path) => deleteWith(core, path),
  };
};

/**
 * The write graph inside any graph, the single package's full client included:
 * its basic-tier members only. The command registry gives every write command
 * this view, so no write command holds an elevated, guest or chat reader. Each
 * member calls the graph's own member on the graph, so a caller's graph whose
 * methods read `this` still works.
 */
const writeGraphOf = (graph: WriteGraph): WriteGraph => ({
  get: (path, extraHeaders) => graph.get(path, extraHeaders),
  getBinary: (path) => graph.getBinary(path),
  fetchUrl: (url) => graph.fetchUrl(url),
  post: (path, body) => graph.post(path, body),
  patch: (path, body) => graph.patch(path, body),
  put: (basePath, body, contentType) => graph.put(basePath, body, contentType),
  delete: (path) => graph.delete(path),
});

export { createWriteGraph, writeGraphOf };
export type { WriteGraph };
