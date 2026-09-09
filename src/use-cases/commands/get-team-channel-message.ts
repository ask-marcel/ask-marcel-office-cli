import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1), messageId: z.string().min(1) });
// A single message honours `$select` only: `$expand=replies` is a 400 on the
// item (probed live 2026-09-09), replies come from
// `list-team-channel-message-replies`.
const inner = buildPickODataListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/messages/${p.messageId}`, baseSchema, ['select']);

const execute: Command['execute'] = async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'Get a single message posted in a channel of a Microsoft Team by its id, through Microsoft Graph on the basic token: the full `chatMessage` with the HTML `body.content`, `from.user`, `createdDateTime`, `attachments[]`, `mentions[]` and `reactions[]`. `--select` trims the projection (`id,from,body`). Its replies are a separate read, `list-team-channel-message-replies`; a reply itself is fetched by its own id the same way, with `replyToId` naming the root post. An unknown message id is named in the error: Graph answers it with `403 Forbidden: UnknownError`, which is not a scope failure.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-get',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Channel message ID, the numeric `id` of a `list-team-channel-messages` entry (a reply id from `list-team-channel-message-replies` works too).',
    },
    ...pickODataOptions(['select']),
  ],
  example: "ask-marcel-office get-team-channel-message --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --message-id '1700000000000'",
  responseShape:
    'single Microsoft Graph `chatMessage` resource: `id`, `replyToId`, `messageType`, `createdDateTime`, `lastModifiedDateTime`, `deletedDateTime`, `subject`, `importance`, `webUrl`, `from { user { id, displayName }, application, device }`, `body { contentType, content }`, `attachments[]`, `mentions[]`, `reactions[]`, `channelIdentity { teamId, channelId }`',
};

export { execute, meta, schema };
