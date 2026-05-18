import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildElevatedListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

// Audit round-8 §1.5: this command was on the basic Teams token throughout
// rounds 6 and 7 (audit had labelled `list-chat-members` "always worked
// without elevation"). Round-8 testing showed the same `Missing scope
// permissions` 403 as the chat-metadata siblings — Graph requires
// `ChatMember.Read` which the basic token doesn't grant. Switch to the
// elevated M365ChatClient identity for parity with `list-chats` /
// `get-chat`.
const baseSchema = z.object({ chatId: z.string().min(1) });
const inner = buildElevatedListCommand((p) => `/chats/${p.chatId}/members`, baseSchema);

// Audit round-7 B3: Graph surfaces the unhelpful `1: NotFound` (the `1:` is
// the Teams thread-id segment, echoed without context) for any missing
// chat-id — empty, malformed, or well-formed-but-unknown. Same rewrite
// shape as round-6's `get-team-channel` fix, but with a chat-id-format hint
// since chat IDs are particularly fiddly (`19:<thread>@thread.v2`).
const execute: Command['execute'] = async (graph, params) => {
  const result = await inner.execute(graph, params);
  if (result.ok) return result;
  if (result.error.type === 'api_error' && /^1:\s*NotFound/i.test(result.error.message)) {
    const chatId = typeof params['chatId'] === 'string' ? params['chatId'] : '<unknown>';
    return err({
      type: 'api_error',
      status: result.error.status,
      message: `NotFound: Microsoft Teams chat not found (chat-id: "${chatId}"). The chat ID format must be \`19:<thread>@thread.v2\` (or \`19:meeting_<id>@thread.v2\` for meeting chats). Source IDs via \`ask-marcel list-chats\` — or URL-decode the \`19%3a...%40thread.v2\` segment of a \`joinUrl\` returned by \`list-calendar-events\`.`,
      code: 'cli_rewrite_chat_not_found',
    });
  }
  return result;
};
const { schema } = inner;

const meta: CommandMeta = {
  summary: 'List the members of a single Microsoft Teams chat.',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/chats/{chat-id}/members',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list-members',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description:
        'Microsoft Teams chat ID, e.g. `19:abc...@thread.v2`. ' +
        'Source the ID via `ask-marcel list-chats` (returns chat metadata for the signed-in user). ' +
        'Alternative sources outside the CLI: the Teams desktop / web client (Open in browser → URL contains the chat thread ID), Microsoft Graph Explorer, ' +
        'or URL-decode the `19%3ameeting_...%40thread.v2` segment of an `onlineMeeting.joinUrl` from `list-calendar-events`.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-chat-members --chat-id '19:abc...@thread.v2'",
  responseShape: 'collection of Microsoft Graph `conversationMember` resources under `value[]`',
  pagination: true,
  needsElevatedToken: true,
};

export { execute, meta, schema };
