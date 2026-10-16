import { z } from 'zod';
import { err } from '@ask-marcel/office-core/domain/result.ts';
import type { ReadCommand, ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { formatZodError } from '@ask-marcel/office-core/use-cases/commands/format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const schema = z.object({ eventId: z.string().min(1) }).extend(odataQuerySchema.shape);

// Slim default `--select` so an LLM doesn't accidentally pull multi-MB
// `contentBytes` for every attachment. `@odata.type` is always returned by
// Graph regardless (and Graph rejects asking for it in `$select`).
const DEFAULT_SELECT = 'id,name,contentType,size,isInline';

const execute: ReadCommand['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const dataWithSelect = parsed.data.select === undefined ? { ...parsed.data, select: DEFAULT_SELECT } : parsed.data;
  const path = appendOData(`/me/events/${parsed.data.eventId}/attachments`, dataWithSelect);
  return graph.get(path);
};

const meta: ReadCommandMeta = {
  summary:
    "List the attachments (file, item, reference) on a single Outlook calendar event. Ships an opinionated default `--select=id,name,contentType,size,isInline` so an LLM doesn't accidentally pull multi-MB `contentBytes` for every attachment. The `@odata.type` discriminator is always returned by Graph regardless of `$select` (and Graph rejects asking for it explicitly). To read one, call `convert-calendar-event-attachment-to-markdown` (or `convert-calendar-event-attachment-to-pdf` for slide / layout fidelity).",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}/attachments',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-list-attachments',
  options: [
    {
      name: 'event-id',
      key: 'eventId',
      required: true,
      description: 'Outlook calendar event ID. Returned by `{read-bin} list-calendar-events` or `get-calendar-event`.',
    },
    ...odataQueryOptions,
  ],
  example: "{read-bin} list-calendar-event-attachments --event-id 'AAMkAGI2...'",
  responseShape:
    'collection of Microsoft Graph `attachment` resources under `value[]` (slim metadata by default — see summary). Graph always includes `@odata.type` and `@odata.mediaContentType` on every entry regardless of `--select` — that discriminator is what the attachment-converting commands branch on.',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Calendars.Read'],
};

export { execute, meta, schema };
