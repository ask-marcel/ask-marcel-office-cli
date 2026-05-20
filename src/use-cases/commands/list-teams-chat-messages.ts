import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Per-chat message history via the Microsoft Teams chat-aggregator. Same
// chatsvcagg-audience bearer as `list-teams-chats-with-messages`. Returns
// the messages of ONE chat — useful when the inlined messages from the
// aggregator entry aren't enough (chatsvcagg typically inlines only the
// last ~10-30 messages).
const schema = z.object({
  chatId: z.string().min(1),
  pageSize: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
  skipToken: z.string().min(1).optional(),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { chatId } = parsed.data;
  const pageSize = parsed.data.pageSize ?? '50';
  const skipToken = parsed.data.skipToken;
  const qs = new URLSearchParams({ pageSize });
  if (skipToken !== undefined) qs.set('skipToken', skipToken);
  return graph.teamsChat(`/api/v2/users/me/chats/${encodeURIComponent(chatId)}/messages?${qs.toString()}`);
};

const meta: CommandMeta = {
  summary:
    "List messages in a single Microsoft Teams chat via the chat-aggregator. Companion to `list-teams-chats-with-messages` when the inlined-per-chat history isn't deep enough. Uses the chatsvcagg-audience bearer captured at login. **Best-effort, may break on Microsoft client updates** — the chatsvcagg surface is not in the public Microsoft Graph API. Pagination via `_skipToken` in the response.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats/{chat-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-list',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description:
        'Teams chat ID — typically `19:<thread>@thread.v2` (1:1 / group) or `19:meeting_<id>@thread.v2` (meeting chat). Source via `list-chats` or `list-teams-chats-with-messages`.',
    },
    { name: 'page-size', key: 'pageSize', required: false, description: 'Number of messages per page (positive integer; default 50). Server may silently cap.' },
    {
      name: 'skip-token',
      key: 'skipToken',
      required: false,
      description: "Opaque pagination cursor returned in the prior response's `_skipToken` field. Omit on the first call.",
    },
  ],
  example: "ask-marcel list-teams-chat-messages --chat-id '19:abc...@thread.v2' --page-size 30",
  responseShape:
    'chatsvcagg envelope: `{ messages: [...], _skipToken?: string }`. Each message has `id`, `from`, `body.content`, `createdDateTime`, etc. **Microsoft-internal schema — fields may change without notice.**',
};

export { execute, meta, schema };
