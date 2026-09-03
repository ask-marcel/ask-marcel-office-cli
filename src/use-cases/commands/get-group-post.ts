import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1), threadId: z.string().min(1), postId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/groups/${p.groupId}/threads/${p.threadId}/posts/${p.postId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Get a single post of a unified (Microsoft 365) group thread by ID, the sibling of `get-mail-message` for a group inbox: the full `post` resource including the HTML `body`. `--select` trims the projection. `--expand attachments` returns the post's attachments inline, each with its base64 `contentBytes`; that is the route to an attachment's bytes, since no separate attachment command exists for posts. Post IDs come from `list-group-thread-posts`; use `convert-group-post-to-markdown` for a readable rendering.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/post-get',
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
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-group-post --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AQMkAD...'",
  responseShape:
    'single Microsoft Graph `post` resource: `id`, `createdDateTime`, `lastModifiedDateTime`, `changeKey`, `categories`, `receivedDateTime`, `hasAttachments`, `body { contentType, content }`, `from` (normally the group address), `sender` (the person who wrote it). With `--expand attachments`, an `attachments[]` array of `fileAttachment` / `itemAttachment` / `referenceAttachment` entries, file attachments carrying `contentBytes` inline.',
};

export { execute, meta, schema };
