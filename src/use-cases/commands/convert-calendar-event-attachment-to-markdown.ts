import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertAttachmentToMarkdown } from './convert-mail-attachment-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  eventId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { eventId, attachmentId } = parsed.data;
  return convertAttachmentToMarkdown(graph, `/me/events/${eventId}/attachments/${attachmentId}`, parsed.data.includeMetadata === 'true');
};

const meta: CommandMeta = {
  summary:
    'Convert an attachment on an Outlook calendar event to markdown. Polymorphic on the attachment’s `@odata.type` (shares the mail-attachment pipeline): fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table, odt/ods/odp via content.xml, plus plain-text passthrough); referenceAttachment resolves via /shares/{token}/driveItem and routes through the same dispatcher; itemAttachment (embedded mail / event / contact) is rendered locally. A pptx deck attached to a meeting yields its speaker notes / slide titles / comments via `## PPTX metadata` (pass `--include-metadata true`). For raw bytes of a pdf/image attachment, save it via `get-mail-attachment` on the equivalent message, or fetch the linked file directly.',
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
  ],
  example: "ask-marcel convert-calendar-event-attachment-to-markdown --event-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` on success (file/reference attachments converted via Graph + turndown; itemAttachment rendered locally). Plain-text source extensions return the raw-bytes envelope; unsupported types return an api_error with status 400.',
  producesBytes: true,
};

export { execute, meta, schema };
