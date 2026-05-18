import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const schema = z.object({ messageId: z.string().min(1) }).extend(odataQuerySchema.shape);

// Default `--select` to a slim metadata set so the LLM never accidentally
// pulls megabytes of `contentBytes`. `@odata.type` is omitted because Graph
// rejects it in `$select` with `Term '@odata.type' is not valid in a $select
// or $expand expression` — discriminator is always returned regardless.
// User-supplied `--select` overrides.
const DEFAULT_SELECT = 'id,name,contentType,size,isInline';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const dataWithSelect = parsed.data.select === undefined ? { ...parsed.data, select: DEFAULT_SELECT } : parsed.data;
  const path = appendOData(`/me/messages/${parsed.data.messageId}/attachments`, dataWithSelect);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    "List the attachments (file, item, reference) on a single Outlook message. The CLI ships an opinionated default `--select=id,name,contentType,size,isInline` so an LLM that doesn't slim the response itself doesn't accidentally pull multi-MB `contentBytes` for every attachment (a single 1.5 MB image attachment would otherwise blow the context window). The `@odata.type` discriminator is always returned by Graph regardless of `$select` (and Graph rejects asking for it explicitly). To fetch the actual bytes, call `get-mail-attachment` for the one you need (or override `--select` if you really want the raw inline payload).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-list-attachments',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `ask-marcel list-mail-messages` or `list-mail-folder-messages`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-mail-attachments --message-id 'AAMkAGI2...'",
  responseShape:
    "collection of Microsoft Graph `attachment` resources under `value[]` (slim metadata by default — see summary). Graph always includes `@odata.type` and `@odata.mediaContentType` on every entry regardless of `--select` — these fields are the discriminator the attachment-converting commands branch on; don't be surprised to see them appear even if you didn't request them.",
  pagination: true,
};

export { execute, meta, schema };
