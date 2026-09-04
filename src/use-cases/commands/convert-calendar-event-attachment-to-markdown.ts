import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertAttachmentToMarkdown } from './convert-mail-attachment-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';
import { keepQuotedOption, keepQuotedSchemaField } from './mail-quote-stripper.ts';
import type { ConversionHints } from './markdown-dispatch.ts';

// The dispatch used to borrow the mail wording here, so an unconvertible event
// attachment told the caller to run `get-mail-attachment --message-id ...`, a
// command that cannot address an event at all. An event has no bytes command of
// its own; expanding the attachments on the event itself is the route to them.
const CALENDAR_HINTS: ConversionHints = {
  pdfNoText:
    'pdf attachment has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). Use `convert-calendar-event-attachment-to-pdf --output-path /tmp/file.pdf` to land the bytes on disk, then read the PDF with a vision-capable model, or run OCR.',
  legacyPpt:
    'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown — there is no pure-JS parser for the format. Use `convert-calendar-event-attachment-to-pdf --output-path /tmp/file.pdf` to render it, then read the PDF with a vision-capable model.',
  image: (ext) =>
    `${ext} attachment is an image and cannot be converted to markdown. Fetch the bytes with \`get-calendar-event --event-id <id> --expand attachments\`, which returns each attachment's base64 \`contentBytes\` inline, and feed them into a vision-capable model. (\`convert-calendar-event-attachment-to-pdf\` is NOT a workaround: Graph's format=pdf rejects images with InputFormatNotSupported.)`,
  generic: (ext) =>
    `${ext} attachment not supported by \`convert-calendar-event-attachment-to-markdown\`. Use \`convert-calendar-event-attachment-to-pdf\` — Graph \`?format=pdf\` accepts 38 input extensions.`,
};

const schema = z.object({
  eventId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
  keepQuoted: keepQuotedSchemaField,
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { eventId, attachmentId } = parsed.data;
  return convertAttachmentToMarkdown(
    graph,
    `/me/events/${eventId}/attachments/${attachmentId}`,
    { includeMetadata: parsed.data.includeMetadata === 'true', keepQuoted: parsed.data.keepQuoted === 'true' },
    CALENDAR_HINTS
  );
};

const meta: CommandMeta = {
  summary:
    'Convert an attachment on an Outlook calendar event to markdown. Polymorphic on the attachment’s `@odata.type` (shares the mail-attachment pipeline): fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, odt/ods/odp via content.xml, plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally. For pptx decks attached to a meeting, `convert-calendar-event-attachment-to-pdf` preserves slide layout (a pptx here yields only its speaker notes / titles / comments via `## PPTX metadata` with `--include-metadata true`). For pdf/rtf/etc. also use the PDF sibling.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'event-id', key: 'eventId', required: true, description: 'Outlook calendar event ID. Returned by `list-calendar-events` or `get-calendar-event`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that event. Returned by `list-calendar-event-attachments`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to surface side-channel content for docx, xlsx, pptx, and OpenDocument attachments. docx → `## DOCX metadata`; xlsx → `## Workbook metadata`; pptx → `## PPTX metadata` (standalone, since pptx has no convertible body); odt/ods/odp → `## OpenDocument metadata`, appended after the converted body. No-op on other attachment types and on itemAttachment renderers.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    keepQuotedOption,
  ],
  example: "ask-marcel-office convert-calendar-event-attachment-to-markdown --event-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` on success (file/reference attachments converted via Graph + turndown; itemAttachment rendered locally). Plain-text source extensions return the raw-bytes envelope; unsupported types return an api_error with status 400.',
  producesBytes: true,
};

export { execute, meta, schema };
