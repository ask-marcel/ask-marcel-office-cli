import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import { wrapNetworkError } from './graph-request.ts';
import { REQUEST_TIMEOUT_MS } from './network-error.ts';

/*
 * The byte-range loop of every upload session: a drive file and an attachment
 * of a mail draft or a To Do task. Each protocol says how a range is sent and
 * which answer completes the file; the loop sends the ranges in order and
 * cancels the session when one fails, so no half-uploaded session is left.
 */

type ChunkedUpload<T> = {
  readonly bytes: Uint8Array;
  readonly chunkSize: number;
  // Sends one range; `contentRange` is `bytes {start}-{end}/{total}`.
  readonly sendChunk: (contentRange: string, chunk: Uint8Array) => Promise<Result<Response, GraphError>>;
  // Whether this answer completes the file (and is not a "send the next range").
  readonly isComplete: (res: Response) => boolean;
  // What the upload answers, read from the response that completed the file.
  readonly readComplete: (res: Response) => Promise<T>;
  readonly cancel: () => Promise<void>;
};

const failed = async <T>(upload: ChunkedUpload<T>, error: GraphError): Promise<Result<never, GraphError>> => {
  await upload.cancel();
  return err(error);
};

/** Sends the bytes range by range; answers what `readComplete` reads from the response that completed the file. */
const uploadChunks = async <T>(upload: ChunkedUpload<T>): Promise<Result<T, GraphError>> => {
  const total = upload.bytes.byteLength;
  for (let start = 0; start < total; start += upload.chunkSize) {
    const end = Math.min(start + upload.chunkSize, total) - 1;
    try {
      const sent = await upload.sendChunk(`bytes ${start}-${end}/${total}`, upload.bytes.slice(start, end + 1));
      if (!sent.ok) return await failed(upload, sent.error);
      if (!sent.value.ok) return await failed(upload, { type: 'api_error', status: sent.value.status, message: `chunk PUT failed at byte ${start}` });
      if (upload.isComplete(sent.value)) return ok(await upload.readComplete(sent.value));
    } catch (e: unknown) {
      return failed(upload, wrapNetworkError(e, 'PUT', `chunk @ byte ${start}`, 'binary'));
    }
  }
  return err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' });
};

// Best-effort cancel of a pre-authenticated upload URL (no bearer); a failure is
// ignored. The DELETE keeps the short budget: a Graph-side cleanup returns promptly.
const cancelUnsigned = async (fetchFn: FetchFn, uploadUrl: string): Promise<void> => {
  try {
    await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
  } catch {
    /* ignore */
  }
};

export { cancelUnsigned, uploadChunks };
export type { ChunkedUpload };
