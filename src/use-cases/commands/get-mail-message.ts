import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

// Audit Alex-session §A: a full Graph `message` resource is 41+ KB by default
// (huge `body.content`, full `internetMessageHeaders`, `uniqueBody`, etc.). LLM
// callers almost always want subject + sender + preview; the body itself is
// usually re-fetched via `get-mail-message-content` only when the preview is
// not enough. Ship a slim default `--select` so the unflagged invocation
// returns ~2-3 KB instead of 41 KB. User `--select foo,bar` always wins.
const DEFAULT_SELECT = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview';

const baseSchema = z.object({ messageId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/me/messages/${p.messageId}`, baseSchema, { defaultSelect: DEFAULT_SELECT });

const meta: CommandMeta = {
  summary:
    "Get a single Outlook message by ID. The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview` so an LLM caller doesn't pull a 41 KB resource just to read a subject line. Pass `--select id,subject,body` (or any other comma-separated field list) to override; for the raw RFC-822 source use `get-mail-message-mime` instead.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `ask-marcel list-mail-messages` or `list-mail-folder-messages`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-mail-message --message-id 'AAMkAGI2...'",
  responseShape:
    'single Microsoft Graph `message` resource projected to the default `--select` set (or, when overridden, to the requested fields). The default omits `body`, `internetMessageHeaders`, and `uniqueBody` — request them explicitly via `--select` when you need the full HTML.',
};

export { execute, meta, schema };
