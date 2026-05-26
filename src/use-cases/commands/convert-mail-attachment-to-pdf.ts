import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { inlineBinary, tagPdfPassthrough } from './fetch-raw-bytes.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';
import { formatZodError } from './format-zod-error.ts';
import { isPdfSource, isPlainTextFilename } from './text-passthrough.ts';

const schema = z.object({
  messageId: z.string().min(1),
  attachmentId: z.string().min(1),
});

const decodeBase64 = (b64: string): Uint8Array => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const safeExtension = (name: string): string => {
  const dot = name.lastIndexOf('.');
  if (dot === -1 || dot === name.length - 1) return 'bin';
  const raw = name.slice(dot + 1).toLowerCase();
  return /^[a-z0-9]{1,8}$/.test(raw) ? raw : 'bin';
};

// Graph's `?format=pdf` rejects image inputs with `InputFormatNotSupported`.
// Audit v1.0.0 §2.4 caught the raw error leaking through. Mirror the markdown
// sibling's friendly guard and point the LLM at `get-mail-attachment` for the
// raw bytes — feeding those into a vision-capable model is the right shape
// for image content anyway.
const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico', 'heic', 'heif']);

const imageHint = (ext: string): string =>
  `${ext} attachment is an image — Graph's format=pdf does not accept image inputs (InputFormatNotSupported). Use \`get-mail-attachment --message-id <id> --attachment-id <id>\` to fetch the bytes (returned base64-encoded) and feed them into a vision-capable model directly; that's the right shape for image content.`;

const extensionOf = (name: string): string => {
  const dot = name.lastIndexOf('.');
  if (dot === -1 || dot === name.length - 1) return '';
  return name.slice(dot + 1).toLowerCase();
};

const convertFileAttachment = async (graph: GraphClient, attachment: { name?: string; contentBytes?: string }): Promise<Result<unknown, GraphError>> => {
  const name = attachment.name ?? 'unnamed';
  const contentBytes = attachment.contentBytes ?? '';
  const bytes = decodeBase64(contentBytes);

  // Plain-text source OR pdf source: skip the upload-convert dance,
  // return raw bytes. PDF is not in Graph's `format=pdf` input list
  // (CDN responds 406 InputFormatNotSupported on pdf → pdf), and the
  // user already has what they want — the raw PDF bytes.
  if (isPlainTextFilename(name) || isPdfSource(name)) {
    return ok({
      contentType: isPdfSource(name) ? 'application/pdf' : 'text/plain',
      size: bytes.byteLength,
      base64: contentBytes,
      note: `pre-checked source (${name}); raw bytes returned without Graph conversion`,
    });
  }

  const lowerExt = extensionOf(name);
  if (IMAGE_EXTENSIONS.has(lowerExt)) return err({ type: 'api_error', status: 415, message: imageHint(lowerExt) });

  // Hardening #2: UUID-only temp file name; never the attacker filename.
  const ext = safeExtension(name);
  const tempName = `${crypto.randomUUID()}.${ext}`;
  const basePath = `/me/drive/root:/.ask-marcel-temp/${tempName}`;

  const uploaded = await graph.put(basePath, bytes, 'application/octet-stream');
  if (!uploaded.ok) return uploaded;
  const uploadedItem = uploaded.value as { id?: string };
  const itemId = uploadedItem.id;
  if (typeof itemId !== 'string') {
    return err({ type: 'api_error', status: 500, message: 'upload returned no driveItem id' });
  }

  const converted = tagPdfPassthrough(await inlineBinary(graph, `/me/drive/items/${itemId}/content?format=pdf`), name);
  // Best-effort cleanup; ignore the err if it fails.
  await graph.delete(`/me/drive/items/${itemId}`);
  // Audit v1.4.0 fresh-pass #7: the temp `.ask-marcel-temp` parent folder
  // used to linger at OneDrive root because we only deleted the file. Now
  // check `--top=1` children; if empty (our file was the last), delete the
  // folder too. Race: a concurrent invocation could upload between the
  // check and the delete — its file would survive (Graph `DELETE` on a
  // path with content does a 412 / 409). Ignore errors either way.
  await cleanupTempFolderIfEmpty(graph);
  return converted;
};

const cleanupTempFolderIfEmpty = async (graph: GraphClient): Promise<void> => {
  const children = await graph.get('/me/drive/root:/.ask-marcel-temp:/children?$top=1&$select=id');
  if (!children.ok) return;
  const body = children.value as { readonly value?: ReadonlyArray<unknown> };
  if ((body.value ?? []).length > 0) return;
  await graph.delete('/me/drive/root:/.ask-marcel-temp');
};

const convertReferenceAttachment = async (graph: GraphClient, attachment: { sourceUrl?: string }): Promise<Result<unknown, GraphError>> => {
  const sourceUrl = attachment.sourceUrl;
  if (typeof sourceUrl !== 'string' || sourceUrl === '') {
    return err({ type: 'api_error', status: 400, message: 'referenceAttachment missing sourceUrl' });
  }
  const resolved = await graph.get(`/shares/${buildShareToken(sourceUrl)}/driveItem`);
  if (!resolved.ok) return resolved;
  const item = resolved.value as { id?: string; name?: string; parentReference?: { driveId?: string } };
  const driveId = item.parentReference?.driveId;
  const itemId = item.id;
  const name = item.name ?? '';
  if (typeof driveId !== 'string' || typeof itemId !== 'string') {
    return err({ type: 'api_error', status: 500, message: 'resolved driveItem missing id or driveId' });
  }
  if (isPlainTextFilename(name) || isPdfSource(name)) {
    return inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content`);
  }
  const refExt = extensionOf(name);
  if (IMAGE_EXTENSIONS.has(refExt)) return err({ type: 'api_error', status: 415, message: imageHint(refExt) });
  return tagPdfPassthrough(await inlineBinary(graph, `/drives/${driveId}/items/${itemId}/content?format=pdf`), name);
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;

  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${attachmentId}`);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;

  // FIX #3: discriminator guard for missing or non-string @odata.type.
  const odataType = a['@odata.type'];
  if (typeof odataType !== 'string') {
    return err({ type: 'api_error', status: 400, message: 'attachment response missing @odata.type discriminator' });
  }

  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return convertFileAttachment(graph, a);
    case '#microsoft.graph.referenceAttachment':
      return convertReferenceAttachment(graph, a);
    case '#microsoft.graph.itemAttachment':
      return err({
        type: 'api_error',
        status: 400,
        message: 'Graph format=pdf cannot accept embedded mail/event/contact items. Use convert-mail-attachment-to-markdown.',
      });
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported attachment type: ${odataType}` });
  }
};

const meta: CommandMeta = {
  summary:
    'Convert an Outlook mail attachment to PDF on the fly. Polymorphic on the attachment’s `@odata.type`: fileAttachment uploads the bytes to a temp folder under /me/drive (large files use Graph’s chunked upload session — no 4 MB ceiling), runs ?format=pdf, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and runs ?format=pdf in place; plain-text source extensions and `pdf` sources short-circuit to a raw-bytes envelope on either path (Graph’s `?format=pdf` does not accept `pdf` as an input format — pdf attachments are returned as-is). itemAttachment (embedded mail/event/contact) is unsupported here — Graph rejects those source types — use convert-mail-attachment-to-markdown instead. Worst-case wall-clock for huge attachments is ~22 minutes (1 metadata GET + up-to-20 chunk PUTs + 1 convert GET + 1 cleanup DELETE, each capped at 60s).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that message. Returned by `list-mail-attachments`.' },
  ],
  example: "ask-marcel convert-mail-attachment-to-pdf --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the PDF bytes, inlined. The CLI follows the SharePoint media-transform redirect internally so the LLM never has to fetch an external URL. Plain-text source extensions and pdf sources short-circuit to `{ contentType, size, base64, note }` with their native bytes; itemAttachment returns api_error 400. Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo` for multi-MB PDFs.',
  producesBytes: true,
};

export { execute, meta, schema };
