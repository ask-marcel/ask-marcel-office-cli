import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Audit v1.0.0 §B2: the previous elevated-token path consistently
// timed out ("elevated token capture timed out — silent SSO ... did not
// yield a Bearer within the deadline"), and the documented `login`
// remediation didn't refresh the M365ChatClient cookies. The regular
// Teams web-client token reaches `/chats/{id}` per Microsoft's chat-get
// docs — the elevation was over-cautious. Switch to the regular token
// to match sibling `list-chat-members`, which has always worked
// without elevation.
const schema = z.object({ chatId: z.string().min(1) });
const { execute } = buildCommand((p) => `/chats/${p.chatId}`, schema);

const meta: CommandMeta = {
  summary:
    'Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). Returns `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. — not the messages (which need `Chat.Read*` and ship as a separate token scope).',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/chats/{chat-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-get',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description: 'Microsoft Teams chat ID, e.g. `19:abc...@thread.v2`. Returned by `list-chats`.',
    },
  ],
  example: "ask-marcel get-chat --chat-id '19:abc...@thread.v2'",
  responseShape: 'single Microsoft Graph `chat` resource',
  scopesRequired: ['Chat.ReadBasic'],
};

export { execute, meta, schema };
