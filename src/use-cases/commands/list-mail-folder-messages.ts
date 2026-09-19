import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { EXCLUDE_MEETING_RESPONSES_OPTION, withMeetingResponseFilter } from './mail-response-filter.ts';
import { odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const baseSchema = z.object({ mailFolderId: z.string().min(1) });
const inner = buildListCommand((p) => `/me/mailFolders/${p.mailFolderId}/messages`, baseSchema);
const schema = z.object({ ...baseSchema.shape, ...odataQuerySchema.shape, excludeMeetingResponses: z.enum(['true', 'false']).optional() });
const execute = withMeetingResponseFilter(schema, inner.execute);

const meta: CommandMeta = {
  summary: 'List the messages inside a specific Outlook mail folder (Inbox, custom folder, etc.).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description:
        'mailFolder ID. Returned by `ask-marcel-office list-mail-folders`. Well-known names also work, e.g. `inbox`, `sentitems`, `drafts`. When listing `drafts`, a `conversationId` `$filter` is not a reliable check for whether a draft already exists on a thread: reply and forward drafts can split across several conversationIds, and `$filter` on Drafts is not read-your-writes consistent. Match client-side on subject and recipients instead, or use the `find-mail-drafts` command, which does exactly that.',
    },
    ...odataQueryOptions,
    EXCLUDE_MEETING_RESPONSES_OPTION,
  ],
  example: "ask-marcel-office list-mail-folder-messages --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
