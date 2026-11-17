import type { AttachmentFile, AttachmentTarget, UploadedAttachment } from '../domain/attachment-upload.ts';
import { ATTACHMENT_MAX_BYTES, SMALL_ATTACHMENT_BYTES } from '../domain/attachment-upload.ts';
import { unlessSendsMail } from '../domain/mail-send-guard.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FetchFn, GraphError, GraphRequestCore } from './graph-request.ts';
import { toBase64 } from './graph-request.ts';
import { REQUEST_TIMEOUT_MS, timeoutMsFor } from './network-error.ts';
import { cancelUnsigned, uploadChunks } from './upload-chunks.ts';

/*
 * The upload of an attachment to a mail draft or a To Do task. The two upload
 * sessions differ (Graph docs, checked 2026-11-17): the Outlook upload URL is
 * pre-authenticated on outlook.office.com and takes no Authorization header;
 * the To Do upload URL is on Graph itself, takes the bearer, and each range
 * goes to `<uploadUrl>/content`. Both take ranges under 4 MB, answer 200 while
 * bytes remain, and 201 with a Location that names the attachment at the end.
 */

type AttachmentUpload = (target: AttachmentTarget, file: AttachmentFile) => Promise<Result<UploadedAttachment, GraphError>>;

// 10 x 320 KiB = 3,276,800 bytes: under the 4 MB that one range may hold.
const CHUNK_BYTES = 10 * 320 * 1024;

// Outlook refuses an upload session for a file under its minimum ("3 MB", in a
// unit the docs do not name); the file then goes in one POST.
const BELOW_SESSION_MINIMUM = 'ErrorAttachmentSizeShouldNotBeLessThanMinimumSize';

const MAIL_UPLOAD_HOSTS: ReadonlySet<string> = new Set(['outlook.office.com', 'outlook.office365.com']);

type Protocol = {
  readonly collection: string;
  readonly fileType: string;
  readonly label: string;
  readonly sessionBody: (item: Readonly<Record<string, unknown>>) => unknown;
  readonly allowsUploadUrl: (url: URL) => boolean;
  readonly sendRange: (uploadUrl: string, headers: Readonly<Record<string, string>>, chunk: Uint8Array) => Promise<Result<Response, GraphError>>;
  readonly cancel: (uploadUrl: string) => Promise<void>;
};

const cancelSigned = async (core: GraphRequestCore, uploadUrl: string): Promise<void> => {
  try {
    await core.send('basic', uploadUrl, { method: 'DELETE', timeoutMs: REQUEST_TIMEOUT_MS });
  } catch {
    /* ignore: a best-effort cleanup */
  }
};

const mailProtocol = (messageId: string, fetchFn: FetchFn): Protocol => ({
  collection: `/me/messages/${messageId}/attachments`,
  fileType: '#microsoft.graph.fileAttachment',
  label: 'an attachment of a mail draft',
  sessionBody: (item) => ({ AttachmentItem: item }),
  allowsUploadUrl: (url) => url.protocol === 'https:' && MAIL_UPLOAD_HOSTS.has(url.host),
  sendRange: async (uploadUrl, headers, chunk) =>
    ok(await fetchFn(uploadUrl, { method: 'PUT', headers, body: chunk as unknown as BodyInit, signal: AbortSignal.timeout(timeoutMsFor('binary')) })),
  cancel: (uploadUrl) => cancelUnsigned(fetchFn, uploadUrl),
});

const todoProtocol = (listId: string, taskId: string, core: GraphRequestCore): Protocol => ({
  collection: `/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
  fileType: '#microsoft.graph.taskFileAttachment',
  label: 'an attachment of a To Do task',
  sessionBody: (item) => ({ attachmentInfo: item }),
  allowsUploadUrl: (url) => url.protocol === 'https:' && url.host === 'graph.microsoft.com',
  sendRange: (uploadUrl, headers, chunk) =>
    core.send('basic', `${uploadUrl}/content`, { method: 'PUT', headers, body: chunk as unknown as BodyInit, timeoutMs: timeoutMsFor('binary') }),
  cancel: (uploadUrl) => cancelSigned(core, uploadUrl),
});

const protocolOf = (target: AttachmentTarget, core: GraphRequestCore, fetchFn: FetchFn): Protocol =>
  target.kind === 'mail-draft' ? mailProtocol(target.messageId, fetchFn) : todoProtocol(target.listId, target.taskId, core);

const contentTypeOf = (file: AttachmentFile): Readonly<Record<string, string>> => (file.contentType === undefined ? {} : { contentType: file.contentType });

const uploaded = (file: AttachmentFile, id: unknown): UploadedAttachment => ({ ...(typeof id === 'string' ? { id } : {}), name: file.name, size: file.bytes.byteLength });

// The id in the Location of the attachment made: `.../Attachments('ID')` (Outlook) or `.../Attachments/ID` (To Do).
const idFromLocation = (location: string | null): string | undefined => {
  const match = /\/attachments(?:\('([^']+)'\)|\/([^/?#]+))/i.exec(location ?? '');
  return match?.[1] ?? match?.[2];
};

const postSmall = async (core: GraphRequestCore, protocol: Protocol, file: AttachmentFile): Promise<Result<UploadedAttachment, GraphError>> => {
  const body = { '@odata.type': protocol.fileType, name: file.name, contentBytes: toBase64(file.bytes), ...contentTypeOf(file) };
  const posted = await core.json('basic', 'POST', protocol.collection, { body });
  return posted.ok ? ok(uploaded(file, (posted.value as { readonly id?: unknown } | undefined)?.id)) : posted;
};

const hostOf = (uploadUrl: string): URL | undefined => {
  try {
    return new URL(uploadUrl);
  } catch {
    return undefined;
  }
};

// The upload URL of a new session, once its host is one this upload may send the file to.
const openSession = async (core: GraphRequestCore, protocol: Protocol, file: AttachmentFile): Promise<Result<string, GraphError>> => {
  const item = { attachmentType: 'file', name: file.name, size: file.bytes.byteLength, ...contentTypeOf(file) };
  const session = await core.json('basic', 'POST', `${protocol.collection}/createUploadSession`, { body: protocol.sessionBody(item) });
  if (!session.ok) return session;
  const uploadUrl = (session.value as { readonly uploadUrl?: unknown } | undefined)?.uploadUrl;
  if (typeof uploadUrl !== 'string') return err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
  const url = hostOf(uploadUrl);
  if (url !== undefined && protocol.allowsUploadUrl(url)) return ok(uploadUrl);
  return err({ type: 'network_error', message: `the upload session URL is on ${url?.host ?? 'no valid host'}, which is not a host this upload sends a file to` });
};

const sendRanges = (protocol: Protocol, uploadUrl: string, bytes: Uint8Array): Promise<Result<string | undefined, GraphError>> =>
  uploadChunks({
    bytes,
    chunkSize: CHUNK_BYTES,
    sendChunk: (contentRange, chunk) => protocol.sendRange(uploadUrl, { 'Content-Range': contentRange, 'Content-Type': 'application/octet-stream' }, chunk),
    isComplete: (res) => res.status === 201,
    readComplete: async (res) => idFromLocation(res.headers.get('location')),
    cancel: () => protocol.cancel(uploadUrl),
  });

const tooLarge = (file: AttachmentFile, maxBytes: number, label: string): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'attachment_too_large',
    message: `${file.name} is ${file.bytes.byteLength} bytes, above the limit of ${maxBytes} bytes for ${label}.`,
  });

const uploadTo = async (core: GraphRequestCore, protocol: Protocol, target: AttachmentTarget, file: AttachmentFile): Promise<Result<UploadedAttachment, GraphError>> => {
  if (file.bytes.byteLength > ATTACHMENT_MAX_BYTES[target.kind]) return tooLarge(file, ATTACHMENT_MAX_BYTES[target.kind], protocol.label);
  if (file.bytes.byteLength < SMALL_ATTACHMENT_BYTES) return postSmall(core, protocol, file);
  const session = await openSession(core, protocol, file);
  if (!session.ok) return session.error.code === BELOW_SESSION_MINIMUM ? postSmall(core, protocol, file) : session;
  const id = await sendRanges(protocol, session.value, file.bytes);
  return id.ok ? ok(uploaded(file, id.value)) : id;
};

/** Attaches a file: in one POST under 3 MB, through an upload session from 3 MB, up to the limit of the target. */
const createAttachmentUpload =
  (core: GraphRequestCore, fetchFn: FetchFn): AttachmentUpload =>
  async (target, file) => {
    const protocol = protocolOf(target, core, fetchFn);
    // An id that resolves the attachments path to a send is refused, as on every write.
    return unlessSendsMail('POST', protocol.collection, () => uploadTo(core, protocol, target, file));
  };

export { createAttachmentUpload };
export type { AttachmentUpload };
