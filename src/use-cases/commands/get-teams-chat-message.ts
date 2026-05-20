import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  chatId: z.string().min(1),
  messageId: z.string().min(1),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { chatId, messageId } = parsed.data;
  return graph.teamsChat(`/api/v2/users/me/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`);
};

const meta: CommandMeta = {
  summary:
    'Return a single Microsoft Teams chat message by its id via the chat-aggregator. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience). **Best-effort, may break on Microsoft client updates** — the chatsvcagg surface is not in the public Microsoft Graph API. Source the chat-id + message-id via `list-teams-chats-with-messages` or `list-teams-chat-messages`.',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats/{chat-id}/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-get',
  options: [
    { name: 'chat-id', key: 'chatId', required: true, description: 'Teams chat ID. Source via `list-chats` or `list-teams-chats-with-messages`.' },
    { name: 'message-id', key: 'messageId', required: true, description: 'Teams chat message ID. Source via `list-teams-chats-with-messages` or `list-teams-chat-messages`.' },
  ],
  example: "ask-marcel get-teams-chat-message --chat-id '19:abc...@thread.v2' --message-id '1700000000000'",
  responseShape: 'single Teams chat message — `id`, `from`, `body.content`, `createdDateTime`, etc. **Microsoft-internal schema — fields may change without notice.**',
};

export { execute, meta, schema };
