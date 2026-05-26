import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Helpers that consolidate the "Graph hands you a 302, follow the CDN
 * redirect, return the bytes" dance used by every command that needs
 * raw file content.
 *
 * Real Graph responses for `/drives/{id}/items/{id}/content` (and the
 * `?format=pdf` / `?format=html` variants) are 302 redirects to a CDN URL.
 * `getBinary` captures the redirect and returns
 * `{ '@microsoft.graph.downloadUrl': '...' }` — NOT inline bytes. To get
 * the bytes we have to follow the URL via `fetchUrl`, which is host-allow-
 * listed for SharePoint / ODSP / Microsoft media-transform domains.
 *
 * Two shapes for the return value depending on what the caller needs:
 *   - `fetchRawBytes(graph, path, opts)` → `Uint8Array` for local
 *     conversion (mammoth / sheetjs / turndown).
 *   - `inlineBinary(graph, path, opts)` → `{ contentType, size, base64 }`
 *     for "stream the bytes back through the JSON envelope" workflows
 *     (PDF conversion, image attachments, anything the LLM consumer
 *     either pipes to disk or saves via --output-path).
 *
 * The historical-version commands pass `elevated: true` so that the
 * initial Graph call is signed with an M365ChatClient token (on the
 * ODSP `logicalPermissions` allow-list); without that, Graph's 302
 * redirects to a streamContent URL whose embedded tempauth is signed
 * by Teams web client identity and rejected by SharePoint with 403.
 */

export type FetchOptions = { readonly elevated?: boolean };

export type InlineBinary = {
  readonly contentType: string;
  readonly size: number;
  readonly base64: string;
};

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const decodeBlobBytes = (blob: Record<string, unknown>): Result<Uint8Array, GraphError> => {
  const b64 = blob['base64'];
  if (typeof b64 === 'string') {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return ok(bytes);
  }
  const text = blob['text'];
  if (typeof text === 'string') {
    return ok(new TextEncoder().encode(text));
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected envelope: response had no @microsoft.graph.downloadUrl, no base64 bytes, and no text body' });
};

const callBinary = (graph: GraphClient, contentPath: string, opts: FetchOptions): Promise<Result<unknown, GraphError>> =>
  opts.elevated ? graph.getBinaryElevated(contentPath) : graph.getBinary(contentPath);

export const fetchRawBytes = async (graph: GraphClient, contentPath: string, opts: FetchOptions = {}): Promise<Result<Uint8Array, GraphError>> => {
  const initial = await callBinary(graph, contentPath, opts);
  if (!initial.ok) return initial;
  const value = initial.value as Record<string, unknown>;

  const downloadUrl = value['@microsoft.graph.downloadUrl'];
  if (typeof downloadUrl === 'string') {
    const followed = await graph.fetchUrl(downloadUrl);
    if (!followed.ok) return followed;
    return decodeBlobBytes(followed.value as Record<string, unknown>);
  }
  return decodeBlobBytes(value);
};

const toInlineBinary = (blob: Record<string, unknown>): Result<InlineBinary, GraphError> => {
  const contentType = typeof blob['contentType'] === 'string' ? blob['contentType'] : 'application/octet-stream';
  const declaredSize = typeof blob['size'] === 'number' ? blob['size'] : undefined;

  const b64 = blob['base64'];
  if (typeof b64 === 'string') {
    return ok({ contentType, size: declaredSize ?? Math.floor((b64.length * 3) / 4), base64: b64 });
  }
  const text = blob['text'];
  if (typeof text === 'string') {
    const bytes = new TextEncoder().encode(text);
    return ok({ contentType, size: declaredSize ?? bytes.byteLength, base64: toBase64(bytes) });
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected envelope: response had no @microsoft.graph.downloadUrl, no base64 bytes, and no text body' });
};

export const inlineBinary = async (graph: GraphClient, contentPath: string, opts: FetchOptions = {}): Promise<Result<InlineBinary, GraphError>> => {
  const initial = await callBinary(graph, contentPath, opts);
  if (!initial.ok) return initial;
  const value = initial.value as Record<string, unknown>;

  const downloadUrl = value['@microsoft.graph.downloadUrl'];
  if (typeof downloadUrl === 'string') {
    const followed = await graph.fetchUrl(downloadUrl);
    if (!followed.ok) return followed;
    return toInlineBinary(followed.value as Record<string, unknown>);
  }
  return toInlineBinary(value);
};

/**
 * Detect Graph's silent-raw-bytes fallback on `?format=pdf` requests.
 *
 * When `format=pdf` succeeds, the response contentType is `application/pdf`.
 * For some inputs (notably historical-version pptx on certain tenants, or
 * reference-attachment edge cases) Graph silently falls back to returning
 * the raw source bytes — same envelope shape, but contentType is the
 * source MIME (or `application/octet-stream`). The audit (round-5 #2)
 * caught this happening on `download-drive-item-version --format pdf` for v79
 * of a pptx: the response said `contentType: "application/octet-stream"`
 * with the exact source byte size, and an LLM that saved it as `.pdf`
 * would have had a corrupt file.
 *
 * Tag the result with `passthrough: true` and a sharp note so the LLM
 * knows the conversion didn't run and saves the bytes with the source
 * extension instead of `.pdf`.
 */
export const tagPdfPassthrough = (result: Result<InlineBinary, GraphError>, sourceLabel: string): Result<unknown, GraphError> => {
  if (!result.ok) return result;
  const ct = result.value.contentType;
  if (ct.startsWith('application/pdf')) return result;
  return ok({
    ...result.value,
    passthrough: true,
    note: `Graph returned \`${ct}\` for ${sourceLabel} — format=pdf conversion was NOT applied (likely no Office Online runtime for this version/format on this tenant). The bytes are the raw source; save with the source extension, not .pdf.`,
  });
};
