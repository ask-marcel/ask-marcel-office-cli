import { z } from 'zod';
import { buildElevatedCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Audit round-8 §1.5: round-6 moved this off elevation on the hypothesis
// that the basic Teams web client token would reach `/chats/{id}`. The
// audit verified the hypothesis was wrong — Graph rejects with `Missing
// scope permissions ... Chat.ReadBasic`. Revert to the elevated
// M365ChatClient path (which DOES carry Chat.ReadBasic).
const schema = z.object({ chatId: z.string().min(1) });
const { execute } = buildElevatedCommand((p) => `/chats/${p.chatId}`, schema);

const meta: CommandMeta = {
  summary:
    'Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). Returns `id`, `topic`, `chatType`, `lastUpdatedDateTime`, etc. — not the messages (which need `Chat.Read*`). Requires the M365ChatClient elevated token captured at login (the basic Teams web client token lacks `Chat.ReadBasic`).',
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
  needsElevatedToken: true,
};

export { execute, meta, schema };
