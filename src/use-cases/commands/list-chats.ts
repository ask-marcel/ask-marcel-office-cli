import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

// Audit v1.0.0 §B2: previously used `buildElevatedListCommand` (M365ChatClient
// token) but the silent-SSO capture consistently timed out and the
// documented `login` remediation didn't refresh the cookies. The Teams
// web-client token reaches `/me/chats` per Microsoft's docs; the
// elevation was over-cautious. Switch to the regular token to match
// sibling `list-chat-members`.
const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/chats', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). Returns chat metadata only — `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. Reading chat *messages* needs the `Chat.Read*` scope which neither token grants.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-chats',
  responseShape: 'collection of Microsoft Graph `chat` resources under `value[]`',
  pagination: true,
  scopesRequired: ['Chat.ReadBasic'],
};

export { execute, meta, schema };
