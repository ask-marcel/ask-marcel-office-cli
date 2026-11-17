import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import { patchOnAllowedRoot } from '../domain/beta-patch.ts';
import type { WriteGraph, WriteOptions } from '../use-cases/ports/graph.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import type { FetchFn, GraphError, GraphRequestCore } from './graph-request.ts';
import { apiErrorFrom, createGraphRequestCore, GRAPH_ROOT, isAllowedFetchUrlHost, wrapNetworkError } from './graph-request.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';
import { cancelUnsigned, uploadChunks } from './upload-chunks.ts';

// Graph beta: reached only by a PATCH on the beta allow-list (domain/beta-patch.ts).
const GRAPH_BETA_ROOT = 'https://graph.microsoft.com/beta';
const SIMPLE_PUT_THRESHOLD = 4 * 1024 * 1024; // 4 MiB
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MiB — Graph requires multiples of 320 KiB; 5 MiB is 16 × 320 KiB

// A write Graph accepted with no body (202 Accepted, 204 No Content) is done.
// Read as JSON, the empty body would report a network error, and a caller could
// send a write that was not idempotent a second time.
const readWriteAnswer = async (res: Response): Promise<unknown> => {
  const text = await res.text();
  return text === '' ? undefined : JSON.parse(text);
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

// 2. PUT chunks to the pre-signed upload URL, no auth header: 200 or 201
// completes the file, 202 asks for the next range.
const putChunks = (fetchFn: FetchFn, uploadUrl: string, body: Uint8Array): Promise<Result<unknown, GraphError>> =>
  uploadChunks({
    bytes: body,
    chunkSize: CHUNK_SIZE,
    sendChunk: async (contentRange, chunk) =>
      ok(
        await fetchFn(uploadUrl, {
          method: 'PUT',
          headers: { 'Content-Range': contentRange },
          body: chunk as unknown as BodyInit,
          signal: AbortSignal.timeout(timeoutMsFor('binary')),
        })
      ),
    isComplete: (res) => res.status === 200 || res.status === 201,
    readComplete: (res): Promise<unknown> => res.json(),
    cancel: () => cancelUnsigned(fetchFn, uploadUrl),
  });

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

// The `If-Match` header of a change that names the ETag it read; none otherwise.
const ifMatchHeader = (options: WriteOptions = {}): Readonly<Record<string, string>> => (options.ifMatch === undefined ? {} : { 'If-Match': options.ifMatch });

const deleteWith = async (core: GraphRequestCore, path: string, options?: WriteOptions): Promise<Result<unknown, GraphError>> => {
  const url = `${GRAPH_ROOT}${path}`;
  try {
    const sent = await core.send('basic', url, { method: 'DELETE', headers: ifMatchHeader(options), timeoutMs: REQUEST_TIMEOUT_MS });
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
    patch: (path, body, options) =>
      patchOnAllowedRoot(path, options?.api, () =>
        core.json('basic', 'PATCH', path, { body, readBody: readWriteAnswer, extraHeaders: ifMatchHeader(options), ...(options?.api === 'beta' ? { root: GRAPH_BETA_ROOT } : {}) })
      ),
    put: (basePath, body, contentType) =>
      body.byteLength <= SIMPLE_PUT_THRESHOLD ? simplePut(core, `${basePath}:/content`, body, contentType) : chunkedPut(core, fetchFn, basePath, body),
    delete: (path, options) => deleteWith(core, path, options),
  };
};

export { createWriteGraph };
export type { WriteGraph };
