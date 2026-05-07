import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ conversationId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/messages?$filter=conversationId eq '${p.conversationId.replace(/'/g, "''")}'`, schema);

const meta: CommandMeta = {
  summary:
    "List every message in a single Outlook conversation (thread) using `$filter=conversationId eq '...'`. Reconstructs a complete thread regardless of which subject lines or folders the replies landed in. Graph rejects combining this filter with `$orderby` (`InefficientFilter` — `conversationId` is not a sortable index), so this command does not order results; the caller can sort by `receivedDateTime` client-side. KQL `$search` does not index `conversationId`, so `$filter` is the only documented Graph idiom for whole-thread retrieval.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: "/me/messages?$filter=conversationId eq '{conversation-id}'",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [
    {
      name: 'conversation-id',
      key: 'conversationId',
      required: true,
      description: 'Outlook `conversationId` of any message in the thread (returned by every mail-listing command and by `get-mail-message`).',
    },
  ],
  example: "ask-marcel list-conversation-messages --conversation-id 'AAQkAD...='",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]` (unordered)',
  pagination: true,
};

export { execute, meta, schema };
