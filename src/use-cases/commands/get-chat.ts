import { z } from 'zod';
import { buildElevatedSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

// round-6 moved this off elevation on the hypothesis
// that the basic Teams web client token would reach `/chats/{id}`. The
// audit verified the hypothesis was wrong — Graph rejects with `Missing
// scope permissions ... Chat.ReadBasic`. Revert to the elevated
// M365ChatClient path (which DOES carry Chat.ReadBasic).
//
// a default `/chats/{id}` response carries `viewpoint`,
// `webUrl`, `tenantId`, `onlineMeetingInfo`, etc. — most callers want only
// chat metadata. Ship a slim default and let the LLM widen via `--select`
// or `--expand members` (which is rejected by Graph at the list level but
// works on single-chat get).
const DEFAULT_SELECT = 'id,topic,chatType,createdDateTime,lastUpdatedDateTime';

const baseSchema = z.object({ chatId: z.string().min(1) });
const { execute, schema } = buildElevatedSelectableCommand((p) => `/chats/${p.chatId}`, baseSchema, { defaultSelect: DEFAULT_SELECT });

const meta: CommandMeta = {
  summary:
    'Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). The CLI ships a slim default `--select=id,topic,chatType,createdDateTime,lastUpdatedDateTime`; pass `--select id,topic,webUrl,onlineMeetingInfo` (or any other comma-separated field list) to widen. Pass `--expand members` to inline membership. Returns metadata only — not the messages (which need `Chat.Read*`). Requires the M365ChatClient elevated token captured at login (the basic Teams web client token lacks `Chat.ReadBasic`).',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/chats/{chat-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-get',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description: 'Microsoft Teams chat ID, e.g. `19:abc...@thread.v2`. Returned by `list-chats`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-chat --chat-id '19:abc...@thread.v2'",
  responseShape:
    'single Microsoft Graph `chat` resource projected to the default `--select` set (or, when overridden, to the requested fields). `--expand members` adds an inline `members[]` array.',
  needsElevatedToken: true,
};

export { execute, meta, schema };
