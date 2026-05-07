import { z } from 'zod';
import { buildElevatedListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildElevatedListCommand(() => '/me/chats', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). Returns chat metadata only — `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. Reading chat *messages* needs the `Chat.Read*` scope which neither token grants. This command requires the elevated M365ChatClient token captured at login.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-chats',
  responseShape: 'collection of Microsoft Graph `chat` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
