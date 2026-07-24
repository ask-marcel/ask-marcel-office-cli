import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildPickODataListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';

// Token: BASIC. `/chats/{id}/members` needs `ChatMember.Read`. Round-8 testing
// found the basic Teams token lacked it (403 `Missing scope permissions`) and
// moved this onto the elevated M365ChatClient identity. As of 2026-06 the basic
// Teams web-client token DOES carry `ChatMember.Read` (`scopes-check` confirms;
// 200s verified live on 1:1 `@unq.gbl.spaces` and meeting `@thread.v2` chats —
// the scope is per-chat-membership, not per-subtype, so group chats ride along),
// so it is back on the basic token via `buildPickODataListCommand`. Basic is preferable: the elevated
// token is login-only and cannot refresh on the command path (see
// `decision_auth_extension_window_login_only`), so depending on it made this
// fail with "Elevated token expired, run login" whenever it was stale.
// NOTE: chat METADATA (`/me/chats`, `/chats/{id}`) still needs `Chat.ReadBasic`,
// which only the elevated token grants — `list-chats` / `get-chat` stay elevated.
//
// `/chats/{id}/members` rejects `$top`, `$orderby`, and
// `$expand` with `BadRequest`. Advertise only the subset Graph honours.
const baseSchema = z.object({ chatId: z.string().min(1) });
const CHAT_MEMBERS_ODATA_KEYS = ['skip', 'select', 'filter'] as const;
const inner = buildPickODataListCommand((p) => `/chats/${p.chatId}/members`, baseSchema, CHAT_MEMBERS_ODATA_KEYS);

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
      message: `NotFound: Microsoft Teams chat not found (chat-id: "${chatId}"). The chat ID format must be \`19:<thread>@thread.v2\` (or \`19:meeting_<id>@thread.v2\` for meeting chats). Source IDs via \`ask-marcel-office list-chats\` — or URL-decode the \`19%3a...%40thread.v2\` segment of a \`joinUrl\` returned by \`list-calendar-events\`.`,
      code: 'cli_rewrite_chat_not_found',
    });
  }
  return result;
};
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the members of a single Microsoft Teams chat. Graph rejects `$top` / `$orderby` / `$expand` on this endpoint, so the CLI advertises only the subset Graph honours (`--skip`, `--select`, `--filter`).',
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
        'Source the ID via `ask-marcel-office list-chats` (returns chat metadata for the signed-in user). ' +
        'Alternative sources outside the CLI: the Teams desktop / web client (Open in browser → URL contains the chat thread ID), Microsoft Graph Explorer, ' +
        'or URL-decode the `19%3ameeting_...%40thread.v2` segment of an `onlineMeeting.joinUrl` from `list-calendar-events`.',
    },
    ...pickODataOptions(CHAT_MEMBERS_ODATA_KEYS),
  ],
  example: "ask-marcel-office list-chat-members --chat-id '19:abc...@thread.v2'",
  responseShape: 'collection of Microsoft Graph `conversationMember` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
