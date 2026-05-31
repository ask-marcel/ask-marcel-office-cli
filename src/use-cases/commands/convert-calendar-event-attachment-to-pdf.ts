import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertAttachmentToPdf } from './convert-mail-attachment-to-pdf.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  eventId: z.string().min(1),
  attachmentId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { eventId, attachmentId } = parsed.data;
  return convertAttachmentToPdf(graph, `/me/events/${eventId}/attachments/${attachmentId}`);
};

const meta: CommandMeta = {
  summary:
    'Convert an attachment on an Outlook calendar event to PDF on the fly (shares the mail-attachment pipeline). fileAttachment uploads the bytes to a temp folder under /me/drive, runs Graph `?format=pdf`, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and converts in place; plain-text and `pdf` sources short-circuit to a raw-bytes envelope (Graph’s `?format=pdf` does not accept `pdf` as an input). image attachments are rejected (Graph rejects image inputs); itemAttachment (embedded mail/event/contact) is unsupported — use convert-calendar-event-attachment-to-markdown. Best for the deck attached to a meeting, where PDF preserves slide layout for a vision-capable LLM.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'event-id', key: 'eventId', required: true, description: 'Outlook calendar event ID. Returned by `list-calendar-events` or `get-calendar-event`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that event. Returned by `list-calendar-event-attachments`.' },
  ],
  example: "ask-marcel convert-calendar-event-attachment-to-pdf --event-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-path ./deck.pdf",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the PDF bytes, inlined. Plain-text and pdf sources short-circuit to `{ contentType, size, base64, note }`; image attachments return api_error 415; itemAttachment returns api_error 400. Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo`.',
  producesBytes: true,
};

export { execute, meta, schema };
