import { z } from 'zod';
import { buildElevatedListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

// Audit round-8 §1.5: round-6 hypothesized that `/me/chats` would succeed
// against the basic Teams web client token; the audit verified it does
// NOT — Graph rejects with `Forbidden: Missing scope permissions ...
// Chat.ReadBasic`. The M365ChatClient elevated identity DOES carry
// Chat.ReadBasic, so revert to the elevated path. If the silent-SSO
// capture times out the command will surface that timeout (documented
// pre-existing failure mode), not the misleading "Missing scope" 403.
const baseSchema = z.object({}).strict();
const { execute, schema } = buildElevatedListCommand(() => '/me/chats', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). Returns chat metadata only — `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. Reading chat *messages* needs `Chat.Read*` which neither token grants. Requires the M365ChatClient elevated token captured at login (the basic Teams web client token lacks `Chat.ReadBasic`).",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-chats',
  responseShape: 'collection of Microsoft Graph `chat` resources under `value[]`',
  pagination: true,
  needsElevatedToken: true,
};

export { execute, meta, schema };
