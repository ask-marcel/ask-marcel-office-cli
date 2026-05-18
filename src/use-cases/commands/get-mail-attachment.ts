import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';

const schema = z.object({ messageId: z.string().min(1), attachmentId: z.string().min(1) }).extend(selectExpandSchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/messages/${parsed.data.messageId}/attachments/${parsed.data.attachmentId}`, parsed.data);
  const result = await graph.get(path);
  if (!result.ok) return result;
  // Audit round-6 §6: surface `contentBytes` as `base64` so the global
  // --output-path interceptor can land the attachment on disk. Keep
  // contentBytes too so existing consumers don't break — the response
  // grows a `base64` mirror of `contentBytes` only for fileAttachment
  // discriminator types (itemAttachment / referenceAttachment have no
  // raw bytes to land on disk).
  const v = result.value as Record<string, unknown>;
  const isFileAttachment = v['@odata.type'] === '#microsoft.graph.fileAttachment';
  const contentBytes = v['contentBytes'];
  if (isFileAttachment && typeof contentBytes === 'string') {
    return ok({ ...v, base64: contentBytes });
  }
  return ok(v);
};

const meta: CommandMeta = {
  summary:
    'Get a single attachment on an Outlook message (metadata, plus the base64 `contentBytes` for file attachments). For fileAttachments, the response also carries a `base64` mirror of `contentBytes` so the global output-path flag can land the bytes on disk in one call. Use `--select id,name,contentType,size` to fetch metadata only and skip the multi-MB `contentBytes` payload.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `ask-marcel list-mail-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID. Returned by `ask-marcel list-mail-attachments`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-mail-attachment --message-id 'AAMkAGI2...' --attachment-id 'AAMkABC...'",
  responseShape:
    'single Microsoft Graph `attachment` resource. fileAttachments include `contentBytes` (Graph) AND `base64` (CLI mirror) so `--output-path` works; itemAttachments and referenceAttachments are returned unchanged.',
  producesBytes: true,
};

export { execute, meta, schema };
