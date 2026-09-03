import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1), threadId: z.string().min(1) });
// Only `$select` and `$expand` reach the server. On this collection Graph
// ignores `$top`, `$skip` and `$orderby` without a word and rejects `$filter`
// outright (`FilteringOnConversations`), probed live 2026-09-03 on a two-post
// thread, so the other passthroughs would promise a slice that never happens.
const { execute, schema } = buildPickODataListCommand((p) => `/groups/${p.groupId}/threads/${p.threadId}/posts`, baseSchema, ['select', 'expand']);

const meta: CommandMeta = {
  summary:
    "List every post in one thread of a unified (Microsoft 365) group inbox: the full `post` resources with the HTML `body.content`, `from`, `sender`, `receivedDateTime` and `hasAttachments`, where `list-group-threads` stops at a truncated `preview`. Graph returns the whole thread in one call with no page cursor, and it silently ignores `$top`, `$skip` and `$orderby` while rejecting `$filter` (probed live 2026-09-03), so only `--select` and `--expand` are exposed; sort on `receivedDateTime` client-side if order matters. `sender` is the person who wrote the post and `from` is normally the group's own address. Access is membership-gated, not scope-gated: a group the signed-in user does not belong to answers `ErrorAccessDenied` even though `list-groups` lists it.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/conversationthread-list-posts',
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
      description: 'Conversation thread ID, the `id` of a `list-group-threads` entry (also inlined by `list-group-conversations --expand threads`).',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office list-group-thread-posts --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...'",
  responseShape:
    'collection of Microsoft Graph `post` resources under `value[]`: `id`, `createdDateTime`, `lastModifiedDateTime`, `changeKey`, `categories`, `receivedDateTime`, `hasAttachments`, `body { contentType, content }`, `from`, `sender`. No `nextLink` is ever emitted. `hasAttachments` is false for a post whose only attachments are inline images. `--expand attachments` inlines every attachment of every post with its base64 `contentBytes`, so use it sparingly.',
};

export { execute, meta, schema };
