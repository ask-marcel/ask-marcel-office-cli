import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { MAIL_MESSAGE_DEFAULT_SELECT } from './mail-message-select.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ messageId: z.string().min(1) });
// Slim default projection shared with list-mail-messages / search-mail-messages
// (see mail-message-select.ts): a full `message` is 41+ KB (body.content,
// internetMessageHeaders, uniqueBody); the slim set is ~2-3 KB. `--select` wins.
const { execute, schema } = buildSelectableCommand((p) => `/me/messages/${p.messageId}`, baseSchema, { defaultSelect: MAIL_MESSAGE_DEFAULT_SELECT });

const meta: CommandMeta = {
  summary:
    "Get a single Outlook message by ID. The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId` (`conversationId` gives you the thread, e.g. for `list-conversation-messages`) so an LLM caller doesn't pull a 41 KB resource just to read a subject line. Pass `--select id,subject,body` (or any other comma-separated field list) to override; for the raw RFC-822 source use `get-mail-message-mime` instead.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description: 'Outlook message ID. Returned by `ask-marcel-office list-mail-messages` or `list-mail-folder-messages`. Accepts `--id` as an alias.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-mail-message --message-id 'AAMkAGI2...'",
  responseShape:
    'single Microsoft Graph `message` resource projected to the default `--select` set (or, when overridden, to the requested fields). The default omits `body`, `internetMessageHeaders`, and `uniqueBody` — request them explicitly via `--select` when you need the full HTML.',
};

export { execute, meta, schema };
