import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ messageId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/messages/${p.messageId}/attachments`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the attachments (file, item, reference) on a single Outlook message.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-list-attachments',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `ask-marcel list-mail-messages` or `list-mail-folder-messages`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-mail-attachments --message-id 'AAMkAGI2...'",
  responseShape: 'collection of Microsoft Graph `attachment` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
