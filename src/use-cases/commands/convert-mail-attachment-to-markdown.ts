import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import {
  embeddedContactToMarkdown,
  embeddedEventToMarkdown,
  embeddedMessageToMarkdown,
  type EmbeddedContact,
  type EmbeddedEvent,
  type EmbeddedMessage,
} from './embedded-item-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { officeToMarkdown } from './office-to-markdown.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';
import { decodeUtf8Text } from './text-passthrough.ts';
import { xlsxToMarkdown } from './xlsx-to-markdown.ts';

const schema = z.object({
  messageId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

const decodeBase64 = (b64: string): Uint8Array => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const PPTX_HINT =
  'pptx attachment not supported by `convert-mail-attachment-to-markdown`. Use `convert-mail-attachment-to-pdf` — Graph PDF conversion preserves slide layout, and a vision-capable LLM reads it more reliably than flattened slide-by-slide bullets. Or pass `--include-metadata true` to extract the side-channel content (speaker notes, comments, hidden slides, properties, tags, links) as a `## PPTX metadata` document.';

// A born-digital PDF attachment's text layer is extracted via unpdf (pdfToMarkdown).
// Only a scanned / image-only PDF (no text layer) falls back to this hint: the bytes
// hold pixels, not text, so a vision model / OCR is the right tool.
const PDF_NO_TEXT_HINT =
  'pdf attachment has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). Use `convert-mail-attachment-to-pdf --output-path /tmp/file.pdf` to land the bytes on disk, then read the PDF with a vision-capable model, or run OCR.';

const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico']);

const imageHint = (ext: string): string =>
  `${ext} attachment is an image and cannot be converted to markdown. Use \`get-mail-attachment --message-id <id> --attachment-id <id>\` to fetch the bytes (returned base64-encoded) and feed them into a vision-capable model directly — that's the right shape for image content. (\`convert-mail-attachment-to-pdf\` is NOT a workaround: Graph's format=pdf rejects images with InputFormatNotSupported.)`;

const genericHint = (ext: string): string =>
  `${ext} attachment not supported by \`convert-mail-attachment-to-markdown\`. Use \`convert-mail-attachment-to-pdf\` — Graph \`?format=pdf\` accepts 38 input extensions.`;

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  if (dot === -1 || dot === filename.length - 1) return '';
  return filename.slice(dot + 1).toLowerCase();
};

const convertFileAttachment = async (attachment: { name?: string; contentBytes?: string }, includeMetadata: boolean): Promise<Result<unknown, GraphError>> => {
  const name = attachment.name ?? 'unnamed';
  const contentBytes = attachment.contentBytes ?? '';
  const bytes = decodeBase64(contentBytes);

  const ext = extensionOf(name);
  if (DOCX_FAMILY.has(ext)) return docxToMarkdown(bytes, { includeMetadata });
  if (XLSX_FAMILY.has(ext)) return xlsxToMarkdown(bytes, { includeMetadata });
  if (PPTX_FAMILY.has(ext)) return includeMetadata ? pptxToMarkdown(bytes) : err({ type: 'api_error', status: 415, message: PPTX_HINT });
  if (ODF_FAMILY.has(ext)) return odfToMarkdown(bytes, { includeMetadata });
  if (ext === 'pdf') return pdfToMarkdown(bytes, PDF_NO_TEXT_HINT);
  if (IMAGE_EXTENSIONS.has(ext)) return err({ type: 'api_error', status: 415, message: imageHint(ext) });

  // Content-sniff: an attachment whose bytes are valid UTF-8 is returned as text
  // (any text file, no extension list); anything else gets the generic hint.
  const text = decodeUtf8Text(bytes);
  if (text !== undefined) return ok({ contentType: 'text/plain', size: bytes.byteLength, text });
  return err({ type: 'api_error', status: 415, message: genericHint(ext === '' ? '<no-extension>' : ext) });
};

const convertReferenceAttachment = async (graph: GraphClient, attachment: { sourceUrl?: string }, includeMetadata: boolean): Promise<Result<unknown, GraphError>> => {
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
  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/content`, name, { includeMetadata });
};

const convertItemAttachment = (attachment: { item?: Record<string, unknown> }): Result<unknown, GraphError> => {
  const item = attachment.item;
  if (!item || typeof item !== 'object') {
    return err({ type: 'api_error', status: 400, message: 'itemAttachment missing inner item' });
  }
  const innerType = item['@odata.type'];
  if (typeof innerType !== 'string') {
    return err({ type: 'api_error', status: 400, message: 'itemAttachment.item missing @odata.type discriminator' });
  }

  // size = UTF-8 byte count (audit §2.1); `md.length` is UTF-16 code units.
  const renderToEnvelope = (md: string): { contentType: 'text/markdown'; size: number; text: string } => ({
    contentType: 'text/markdown',
    size: new TextEncoder().encode(md).byteLength,
    text: md,
  });

  switch (innerType) {
    case '#microsoft.graph.message':
      return ok(renderToEnvelope(embeddedMessageToMarkdown(item as EmbeddedMessage)));
    case '#microsoft.graph.event':
      return ok(renderToEnvelope(embeddedEventToMarkdown(item as EmbeddedEvent)));
    case '#microsoft.graph.contact':
      return ok(renderToEnvelope(embeddedContactToMarkdown(item as EmbeddedContact)));
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported embedded item type: ${innerType}` });
  }
};

// Fetch an attachment by its full Graph path and convert it to markdown,
// branching on the polymorphic `@odata.type`. Path-agnostic so both the mail
// (`/me/messages/{id}/attachments/{id}`) and calendar-event
// (`/me/events/{id}/attachments/{id}`) commands share one implementation.
const convertAttachmentToMarkdown = async (graph: GraphClient, attachmentPath: string, includeMetadata: boolean): Promise<Result<unknown, GraphError>> => {
  const fetched = await graph.get(attachmentPath);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;

  const odataType = a['@odata.type'];
  if (typeof odataType !== 'string') {
    return err({ type: 'api_error', status: 400, message: 'attachment response missing @odata.type discriminator' });
  }

  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return convertFileAttachment(a, includeMetadata);
    case '#microsoft.graph.referenceAttachment':
      return convertReferenceAttachment(graph, a, includeMetadata);
    case '#microsoft.graph.itemAttachment':
      return convertItemAttachment(a);
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported attachment type: ${odataType}` });
  }
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;
  return convertAttachmentToMarkdown(graph, `/me/messages/${messageId}/attachments/${attachmentId}`, parsed.data.includeMetadata === 'true');
};

const meta: CommandMeta = {
  summary:
    'Convert an Outlook mail attachment to markdown. Polymorphic on the attachment’s `@odata.type`: fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, odt/ods/odp via content.xml, pdf via text-layer extraction (unpdf → text/plain), plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally via dedicated renderers. For pptx attachments, `convert-mail-attachment-to-pdf` is recommended (Graph PDF preserves slide layout). A scanned / image-only PDF (no text layer) and rtf/etc. point to the PDF sibling. Loop/Fluid/Whiteboard reference-attachments use Graph `?format=html` (the four inputs Microsoft documents).',
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
        'Pass `--include-metadata true` to surface side-channel content for docx, xlsx, pptx, and OpenDocument attachments (file + reference). docx → `## DOCX metadata` (properties, people, hyperlinks, comments, tracked changes, hidden text, fields, bookmarks); xlsx → `## Workbook metadata` (properties, external relationships, defined names, hidden / very-hidden sheets, cell + threaded comments, persons); pptx → `## PPTX metadata` (properties, external relationships, slide tags, comment authors + comments, per-slide title / speaker notes / hidden flag) as a standalone document, since pptx has no convertible body; odt/ods/odp → `## OpenDocument metadata` (Dublin Core + ODF properties, keywords, user-defined fields), appended after the converted body. Each OOXML family also covers its macro-enabled and template variants, with a `### Macros (VBA)` section flagging an embedded `vbaProject.bin`. No-op on other attachment types and on itemAttachment renderers.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel convert-mail-attachment-to-markdown --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` on success (file/reference attachments converted via Graph + turndown; itemAttachment rendered locally). Plain-text source extensions return the raw-bytes envelope; unsupported types return an api_error with status 400.',
  producesBytes: true,
};

export { convertAttachmentToMarkdown, execute, meta, schema };
