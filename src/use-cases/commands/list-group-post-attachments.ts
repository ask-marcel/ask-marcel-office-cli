import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1), threadId: z.string().min(1), postId: z.string().min(1) });

// Slim default `--select`, matching the mail and calendar siblings: without it
// Graph returns every attachment's base64 `contentBytes`, which is the whole
// reason this command exists next to `get-group-post --expand attachments`.
const DEFAULT_SELECT = 'id,name,contentType,size,isInline';

// Only `$select` and `$expand` reach the server. Probed live 2026-09-03 on a
// post with 7 attachments: `$top=1` returned all 7, `$skip=1` returned all 7,
// `$orderby` changed nothing, and `$filter=isInline eq false` returned all 7
// although every one of them is inline. Graph ignores them without a word,
// which is worse than refusing them, so they are not advertised.
const { execute, schema } = buildPickODataListCommand((p) => `/groups/${p.groupId}/threads/${p.threadId}/posts/${p.postId}/attachments`, baseSchema, ['select', 'expand'], {
  defaultSelect: DEFAULT_SELECT,
});

const meta: CommandMeta = {
  summary:
    'List the attachments (file, item, reference) on one post of a unified (Microsoft 365) group thread. Ships the slim default `--select=id,name,contentType,size,isInline` the mail and calendar siblings use, so a caller sees what is attached without pulling any bytes — the staged alternative to `get-group-post --expand attachments`, which inlines EVERY attachment at once and times out on a post carrying a multi-MB file. Graph returns the whole collection in one response and silently ignores `$top`, `$skip`, `$orderby` and `$filter` (probed live 2026-09-03), so only `--select` and `--expand` are exposed. A post whose only attachments are inline images reports `hasAttachments: false`, so call this whenever the body shows `cid:` references. Read one with `convert-group-post-attachment-to-markdown`, or fetch its bytes with `get-group-post-attachment`.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/post-list-attachments',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group you belong to.',
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
  example: "ask-marcel-office list-group-post-attachments --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AQMkAD...'",
  responseShape:
    'collection of Microsoft Graph `attachment` resources under `value[]` (slim metadata by default — see summary), with no page cursor. Graph always includes `@odata.type` and `@odata.mediaContentType` on every entry regardless of `--select`; that discriminator is what the converting sibling branches on. An inline image carries `isInline: true` and a `contentId` matching a `cid:` reference in the post body.',
};

export { execute, meta, schema };
