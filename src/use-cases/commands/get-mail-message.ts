import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ messageId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/me/messages/${p.messageId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get a single Outlook message by ID, including subject, sender, body, and flags. Pass `--select id,subject,from,receivedDateTime` to fetch only the fields the LLM needs (a full message body can be 50+ KB; the audit found this swelling LLM context unnecessarily).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `ask-marcel list-mail-messages` or `list-mail-folder-messages`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-mail-message --message-id 'AAMkAGI2...' --select id,subject,from,receivedDateTime",
  responseShape: 'single Microsoft Graph `message` resource (or projection of the requested `--select` fields)',
};

export { execute, meta, schema };
