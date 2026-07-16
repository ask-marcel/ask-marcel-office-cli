import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertFetchedAttachment } from './convert-mail-attachment-to-markdown.ts';
import { base64ToBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { convertZipArchive } from './zip-archive-to-markdown.ts';

const schema = z.object({
  messageId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

// A fileAttachment whose bytes are a zip archive — by extension or content-type.
// (`x-zip-compressed` is the legacy Windows/Outlook content-type for .zip.)
const isZipAttachment = (a: Record<string, unknown>): boolean => {
  if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') return false;
  const name = typeof a['name'] === 'string' ? a['name'].toLowerCase() : '';
  const contentType = typeof a['contentType'] === 'string' ? a['contentType'].toLowerCase() : '';
  return name.endsWith('.zip') || contentType === 'application/zip' || contentType === 'application/x-zip-compressed';
};

// Content-types whose bytes are a convertible Office/text document. Senders and
// Outlook mislabel attachments (a real spreadsheet saved as `report.jpg`), so when
// the filename extension would misroute but the content-type is one of these, the
// polymorphic reader routes by the content-type instead. Generic types
// (octet-stream) are absent, so a correctly-named file with a vague content-type
// still routes by its name (matching the zip-by-extension behaviour above).
const CONTENT_TYPE_EXTENSIONS: Readonly<Record<string, string>> = {
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
  'application/vnd.oasis.opendocument.spreadsheet': 'ods',
  'application/vnd.oasis.opendocument.text': 'odt',
  'application/vnd.oasis.opendocument.presentation': 'odp',
  'application/pdf': 'pdf',
  'application/msword': 'doc',
  'application/vnd.ms-excel': 'xls',
  'text/csv': 'csv',
};

// Rename a fileAttachment to the extension its content-type implies when the
// filename's own extension disagrees, so the shared dispatch picks the right
// converter. Non-file attachments and already-correct names pass through unchanged.
const nameByContentType = (a: Record<string, unknown>): Record<string, unknown> => {
  if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') return a;
  const contentType = typeof a['contentType'] === 'string' ? a['contentType'].toLowerCase() : '';
  const ext = Object.hasOwn(CONTENT_TYPE_EXTENSIONS, contentType) ? CONTENT_TYPE_EXTENSIONS[contentType] : undefined;
  if (ext === undefined) return a;
  const name = typeof a['name'] === 'string' ? a['name'] : '';
  const dot = name.lastIndexOf('.');
  if ((dot === -1 ? '' : name.slice(dot + 1).toLowerCase()) === ext) return a;
  return { ...a, name: `${dot === -1 ? name || 'attachment' : name.slice(0, dot)}.${ext}` };
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';

  // Single fetch, then auto-route by content-type: a zip is unpacked + converted
  // entry-by-entry; everything else (docx/xlsx/pptx/odf/csv/pdf/.msg/legacy/text,
  // referenceAttachment, embedded item) goes through the markdown dispatch — so a
  // caller that just wants to "read this attachment" never has to pick the command.
  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${attachmentId}`);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;

  if (isZipAttachment(a)) {
    const contentBytes = a['contentBytes'];
    if (typeof contentBytes !== 'string') {
      return err({ type: 'api_error', status: 400, message: 'zip fileAttachment has no contentBytes to unpack (the attachment may be empty).' });
    }
    return convertZipArchive(base64ToBytes(contentBytes), includeMetadata);
  }
  return convertFetchedAttachment(graph, nameByContentType(a), includeMetadata);
};

const meta: CommandMeta = {
  summary:
    'Read an Outlook mail attachment whatever it is — one command that auto-routes by file type, preferring the content-type when the filename extension is misleading (a real `.jpg` that is actually a spreadsheet still converts), so a caller never has to choose between the convert-mail-attachment-* siblings. A `.zip` fileAttachment is unpacked and every entry converted (mirrors `convert-mail-attachment-zip-to-markdown`, returning the `{ count, files }` envelope; legacy GBK/CP437 names decoded). Any other attachment — docx/xlsx/pptx/odt/ods/odp + macro/template variants → markdown, csv → table, pdf → text layer (with `pageCount`), legacy .xls/.doc extracted, an inner Outlook .msg rendered recursively, plain text passed through, referenceAttachment resolved via `/shares`, and itemAttachment (embedded mail/event/contact) rendered — goes through the same dispatch as `convert-mail-attachment-to-markdown` (returning its `{ contentType, size, text }` envelope). Images, scanned/image-only PDFs, and legacy .ppt return an actionable 415 pointing at `convert-mail-attachment-to-pdf` + a vision model or `get-mail-attachment` for the raw bytes. Pass `--include-metadata true` to append Office side-channel metadata. Use the explicit `convert-mail-attachment-to-markdown` / `-to-pdf` / `-zip` siblings only when you need to force a specific output format.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that message. Returned by `list-mail-attachments`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append each converted Office file’s side-channel metadata block (docx / xlsx / pptx / OpenDocument). No-op on images, embedded items, and plain text.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel-office read-mail-attachment --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    'Polymorphic by attachment content-type. A zip → `{ count, files: [{ path, contentType, size, text } | { path, note }], truncated? }` (the convert-mail-attachment-zip-to-markdown shape). Everything else → `{ contentType: "text/markdown" | "text/plain", size, text, pageCount? }` (the convert-mail-attachment-to-markdown shape; `pageCount` present for PDF sources). Unsupported types (image / scanned PDF / legacy .ppt) return an api_error (415/400) naming the right next command.',
  producesBytes: true,
};

export { execute, meta, schema };
