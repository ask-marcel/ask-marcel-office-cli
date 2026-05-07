import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/messages', baseSchema);

const meta: CommandMeta = {
  summary:
    'List the most recent messages from across the signed-in user’s entire Outlook mailbox (every folder including Sent, Archive, Junk; default sort `receivedDateTime` desc). Use `list-mail-folder-messages` to scope to a single folder such as Inbox.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-mail-messages',
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
