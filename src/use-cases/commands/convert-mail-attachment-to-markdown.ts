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
import { officeToMarkdown } from './office-to-markdown.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';
import { isPlainTextFilename } from './text-passthrough.ts';
import { xlsxToMarkdown } from './xlsx-to-markdown.ts';

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

const PPTX_HINT =
  'pptx attachment not supported by `convert-mail-attachment-to-markdown`. Use `convert-mail-attachment-to-pdf` — Graph PDF conversion preserves slide layout, and a vision-capable LLM reads it more reliably than flattened slide-by-slide bullets.';

const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico']);

const imageHint = (ext: string): string =>
  `${ext} attachment is an image and cannot be converted to markdown. Use \`get-mail-attachment --message-id <id> --attachment-id <id>\` to fetch the bytes (returned base64-encoded). \`convert-mail-attachment-to-pdf\` will also succeed and wrap the same image in a PDF container, but that is rarely useful — feed the raw bytes into a vision-capable model instead.`;

const genericHint = (ext: string): string =>
  `${ext} attachment not supported by \`convert-mail-attachment-to-markdown\`. Use \`convert-mail-attachment-to-pdf\` — Graph \`?format=pdf\` accepts 38 input extensions.`;

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  if (dot === -1 || dot === filename.length - 1) return '';
  return filename.slice(dot + 1).toLowerCase();
};

const convertFileAttachment = async (attachment: { name?: string; contentBytes?: string }): Promise<Result<unknown, GraphError>> => {
  const name = attachment.name ?? 'unnamed';
  const contentBytes = attachment.contentBytes ?? '';
  const bytes = decodeBase64(contentBytes);

  if (isPlainTextFilename(name)) {
    return ok({
      contentType: 'text/plain',
      size: bytes.byteLength,
      base64: contentBytes,
      note: `pre-checked plain-text source (${name}); raw bytes returned without conversion`,
    });
  }

  const ext = extensionOf(name);
  if (ext === 'docx') return docxToMarkdown(bytes);
  if (ext === 'xlsx') return xlsxToMarkdown(bytes);
  if (ext === 'pptx') return err({ type: 'api_error', status: 415, message: PPTX_HINT });
  if (IMAGE_EXTENSIONS.has(ext)) return err({ type: 'api_error', status: 415, message: imageHint(ext) });
  return err({ type: 'api_error', status: 415, message: genericHint(ext === '' ? '<no-extension>' : ext) });
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
  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/content`, name);
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

  const renderToEnvelope = (md: string): { contentType: 'text/markdown'; size: number; text: string } => ({
    contentType: 'text/markdown',
    size: md.length,
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

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;

  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${attachmentId}`);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;

  const odataType = a['@odata.type'];
  if (typeof odataType !== 'string') {
    return err({ type: 'api_error', status: 400, message: 'attachment response missing @odata.type discriminator' });
  }

  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return convertFileAttachment(a as { name?: string; contentBytes?: string });
    case '#microsoft.graph.referenceAttachment':
      return convertReferenceAttachment(graph, a as { sourceUrl?: string });
    case '#microsoft.graph.itemAttachment':
      return convertItemAttachment(a as { item?: Record<string, unknown> });
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported attachment type: ${odataType}` });
  }
};

const meta: CommandMeta = {
  summary:
    'Convert an Outlook mail attachment to markdown. Polymorphic on the attachment’s `@odata.type`: fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally via dedicated renderers. For pptx attachments, `convert-mail-attachment-to-pdf` is recommended (Graph PDF preserves slide layout). For pdf/rtf/odt/etc. also use the PDF sibling. Loop/Fluid/Whiteboard reference-attachments use Graph `?format=html` (the four inputs Microsoft documents).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that message. Returned by `list-mail-attachments`.' },
  ],
  example: "ask-marcel convert-mail-attachment-to-markdown --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` on success (file/reference attachments converted via Graph + turndown; itemAttachment rendered locally). Plain-text source extensions return the raw-bytes envelope; unsupported types return an api_error with status 400.',
};

export { execute, meta, schema };
