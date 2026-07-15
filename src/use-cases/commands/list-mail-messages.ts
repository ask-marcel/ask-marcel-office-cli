import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { MAIL_MESSAGE_DEFAULT_SELECT } from './mail-message-select.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
// Slim default projection shared with search-mail-messages / get-mail-message
// (see mail-message-select.ts); at 25 messages/page the full Graph projection
// runs ~1 MB vs ~30-60 KB slim. A user-supplied `--select` always wins.
const { execute, schema } = buildListCommand(() => '/me/messages', baseSchema, { defaultSelect: MAIL_MESSAGE_DEFAULT_SELECT });

const meta: CommandMeta = {
  summary:
    "List the most recent messages from across the signed-in user's entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId` (`conversationId` groups messages into a thread and can be handed to `list-conversation-messages`) so a page of 25 messages stays ~30-60 KB instead of ~1 MB. Pass `--select id,subject,body` (or any other comma-separated field list) to override. Use `list-mail-folder-messages` to scope to a single folder such as Inbox.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [...odataQueryOptions],
  example: 'ask-marcel-office list-mail-messages',
  responseShape:
    'collection of Microsoft Graph `message` resources under `value[]`, each projected to the default `--select` set (or the requested fields when overridden). The default omits `body`, `internetMessageHeaders`, and `uniqueBody`.',
  pagination: true,
};

export { execute, meta, schema };
