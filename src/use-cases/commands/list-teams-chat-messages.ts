import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Per-chat message history via the Microsoft Teams chat substrate. Same
// chatsvcagg-audience bearer as `list-teams-chats-with-messages`. Returns
// up to the 200 most recent messages in a single response.
//
// No pagination. The chatsvcagg `/api/v1/chats/{id}/messages` route caps
// every request at 200 messages and ignores `pageSize` AND every cursor
// parameter we've probed (`messageToken`, `syncState`, `continuationToken`,
// `pageToken`, `cursor`, etc.) — verified empirically on 2026-05-21. Teams
// web itself uses WebSockets / SignalR for scrollback, NOT this HTTP route,
// so the route was never wired for paginated callers. The official
// `Chat.Read` Graph scope would unlock the paginated public-API endpoint
// (`/v1.0/chats/{id}/messages`) but is locked behind the Teams web client
// `appid`'s fixed scope ceiling. See `gotcha_chatsvcagg_substrate_moved`
// in project memory for the substrate audit.
const schema = z.object({
  chatId: z.string().min(1),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { chatId } = parsed.data;
  return graph.teamsChat(`/api/v1/chats/${encodeURIComponent(chatId)}/messages`);
};

const meta: CommandMeta = {
  summary:
    "List the most recent messages in a single Microsoft Teams chat via the chat substrate. Companion to `list-teams-chats-with-messages` when the inlined `lastMessage` isn't deep enough. Uses the chatsvcagg-audience bearer captured at login. **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsoft Graph API. **No pagination**: the route caps at the 200 most recent messages per chat and the CLI cannot reach older history (Teams web itself uses WebSockets for scrollback, and the official `Chat.Read` Graph scope that would enable paginated reads is outside the appid's scope ceiling).",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/csa/{region}/api/v1/chats/{chat-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-list',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description: 'Teams chat ID — typically `19:<thread>@unq.gbl.spaces` (1:1) or `19:<thread>@thread.v2` (group). Source via `list-chats` or `list-teams-chats-with-messages`.',
    },
  ],
  example: "ask-marcel list-teams-chat-messages --chat-id '19:abc...@unq.gbl.spaces'",
  responseShape:
    'Substrate envelope: `{ messages: [...], messageToken: string }`. Returns up to the 200 most recent messages per chat — older history is NOT reachable via this endpoint. Each message has `id`, `from`, `imDisplayName`, `content`, `contentType`, `composeTime`, `originalArrivalTime`, `sequenceId`, etc. `messageToken` is returned for forward compatibility but is currently a static snapshot identifier (server ignores it as a pagination cursor). **Microsoft-internal schema — fields may change without notice.** For history older than the 200 most recent, use `list-teams-chat-history` (rides the IC3 substrate with a working syncState cursor).',
  stability: 'experimental',
};

export { execute, meta, schema };
