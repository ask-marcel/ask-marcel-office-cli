import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

// Audit Alex-session §A: matches `get-mail-message`'s slim default — at 25
// messages per page, the full Graph projection runs ~1 MB; the slim default
// is ~30-60 KB. User `--select foo,bar` always wins.
const DEFAULT_SELECT = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/messages', baseSchema, { defaultSelect: DEFAULT_SELECT });

const meta: CommandMeta = {
  summary:
    "List the most recent messages from across the signed-in user's entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview` so a page of 25 messages stays ~30-60 KB instead of ~1 MB. Pass `--select id,subject,body` (or any other comma-separated field list) to override. Use `list-mail-folder-messages` to scope to a single folder such as Inbox.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-mail-messages',
  responseShape:
    'collection of Microsoft Graph `message` resources under `value[]`, each projected to the default `--select` set (or the requested fields when overridden). The default omits `body`, `internetMessageHeaders`, and `uniqueBody`.',
  pagination: true,
};

export { execute, meta, schema };
