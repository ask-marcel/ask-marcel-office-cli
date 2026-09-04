import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';

const schema = z.object({ groupId: z.string().min(1), threadId: z.string().min(1), postId: z.string().min(1), attachmentId: z.string().min(1) }).extend(selectExpandSchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId, attachmentId } = parsed.data;
  const path = appendOData(`/groups/${groupId}/threads/${threadId}/posts/${postId}/attachments/${attachmentId}`, parsed.data);
  const result = await graph.get(path);
  if (!result.ok) return result;
  // `base64` mirrors `contentBytes` so the global --output-path interceptor can
  // land the file on disk, exactly as `get-mail-attachment` does. Only a
  // fileAttachment has raw bytes; the other two subtypes are returned unchanged.
  const value = result.value as Record<string, unknown>;
  const contentBytes = value['contentBytes'];
  if (value['@odata.type'] === '#microsoft.graph.fileAttachment' && typeof contentBytes === 'string') {
    return ok({ ...value, base64: contentBytes });
  }
  return ok(value);
};

const meta: CommandMeta = {
  summary:
    'Get a single attachment on one post of a unified (Microsoft 365) group thread, the `get-mail-attachment` sibling for a group inbox. Prefer it over `get-group-post --expand attachments`, which expands every attachment at once. fileAttachments carry a `base64` mirror of `contentBytes` so the global output-path flag lands the bytes on disk in one call; with an output-path set both byte fields are stripped from stdout in favour of `savedTo`. Pass `--select id,name,contentType,size` for metadata only. This is also the route to an image attached to a post: fetch the bytes and feed them to a vision-capable model.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group the signed-in user belongs to.',
    },
    {
      name: 'thread-id',
      key: 'threadId',
      required: true,
      description: 'Conversation thread ID, the `id` of a `list-group-threads` entry.',
    },
    {
      name: 'post-id',
      key: 'postId',
      required: true,
      description: 'Post ID inside that thread. Returned by `list-group-thread-posts`.',
    },
    {
      name: 'attachment-id',
      key: 'attachmentId',
      required: true,
      description: 'Attachment ID inside that post. Returned by `list-group-post-attachments`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-group-post-attachment --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AQMkAD...' --attachment-id 'AAMkAD...'",
  responseShape:
    'single Microsoft Graph `attachment` resource. fileAttachments include `contentBytes` (Graph) AND `base64` (CLI mirror) so `--output-path` works; with `--output-path` set, both byte fields are stripped from stdout and replaced by `savedTo`. itemAttachments and referenceAttachments are returned unchanged.',
  producesBytes: true,
};

export { execute, meta, schema };
