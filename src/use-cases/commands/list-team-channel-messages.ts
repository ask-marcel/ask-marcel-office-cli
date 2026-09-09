import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import { CHANNEL_MESSAGES_TOP_OPTION, withChannelMessagesTopCap } from './channel-message-page.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
// Probed live 2026-09-09 on the basic token: `$top` (max 50), `$select` and
// `$expand=replies` reach the server; `$filter` and `$orderby` are 400s and
// `$skip` is undocumented, so paging is the `@odata.nextLink` Graph returns.
const inner = buildPickODataListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/messages`, baseSchema, ['top', 'select', 'expand']);

const execute: Command['execute'] = withChannelMessagesTopCap(async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params)));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the messages posted in a single channel of a Microsoft Team, newest first, through Microsoft Graph on the basic token (no chat-substrate warm-up, unlike the Teams chat commands): the `chatMessage` resources with the HTML `body.content`, `from.user` (`displayName`, `id`), `createdDateTime`, `messageType` (`message` for a post, `systemEventMessage` for a membership or channel event), `attachments[]` as references (a shared file carries `contentUrl` and `name`), `mentions[]` and `reactions[]`. Each entry is a ROOT post: replies are not inlined unless `--expand replies` is passed, or fetched with `list-team-channel-message-replies`. A deleted message stays in the list with `deletedDateTime` set and an empty body. `--top` pages up to 50 at a time; older history continues through the `next:` footer with `next-page`. Find the ids with `list-joined-teams` then `list-team-channels`; an unknown or malformed channel id is named in the error instead of the bare `UnknownError` Graph answers.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/channel-list-messages',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    {
      name: 'channel-id',
      key: 'channelId',
      required: true,
      description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`, or `get-team-primary-channel` for General.',
    },
    CHANNEL_MESSAGES_TOP_OPTION,
    ...pickODataOptions(['select']),
    {
      name: 'expand',
      key: 'expand',
      required: false,
      description: "OData $expand: pass `replies` to inline each root post's replies (the only navigation Graph expands on this collection). Increases response size.",
    },
  ],
  example: "ask-marcel-office list-team-channel-messages --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --top 20",
  responseShape:
    'collection of Microsoft Graph `chatMessage` resources under `value[]`: `id`, `replyToId` (null on a root post), `messageType`, `createdDateTime`, `lastModifiedDateTime`, `deletedDateTime`, `subject`, `importance`, `webUrl`, `from { user { id, displayName }, application, device }`, `body { contentType, content }`, `attachments[]`, `mentions[]`, `reactions[]`, `channelIdentity { teamId, channelId }`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
