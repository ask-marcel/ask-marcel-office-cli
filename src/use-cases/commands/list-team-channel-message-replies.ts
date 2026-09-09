import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import { CHANNEL_MESSAGES_TOP_OPTION, withChannelMessagesTopCap } from './channel-message-page.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1), messageId: z.string().min(1) });
// Same OData surface as the channel's message list, probed live 2026-09-09:
// `$top` (max 50) and `$select`; paging by `@odata.nextLink`. An unknown
// message id answers 200 with an empty page rather than an error, so an empty
// list is not proof the post exists.
const inner = buildPickODataListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/messages/${p.messageId}/replies`, baseSchema, ['top', 'select']);

const execute: Command['execute'] = withChannelMessagesTopCap(async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params)));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the replies to one root post in a channel of a Microsoft Team, newest first, through Microsoft Graph on the basic token: the same `chatMessage` shape as `list-team-channel-messages`, each with `replyToId` set to the root post. This is the thread under a post; `list-team-channel-messages` can inline the same replies on every root post of a page instead, through its `expand` option. `--top` pages up to 50 at a time, older replies continue through the `next:` footer with `next-page`. Graph answers an unknown message id with an EMPTY list rather than an error, so an empty result does not prove the post exists: `get-team-channel-message` does.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-list-replies',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    { name: 'message-id', key: 'messageId', required: true, description: 'Root post ID, the numeric `id` of a `list-team-channel-messages` entry whose `replyToId` is null.' },
    CHANNEL_MESSAGES_TOP_OPTION,
    ...pickODataOptions(['select']),
  ],
  example: "ask-marcel-office list-team-channel-message-replies --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --message-id '1700000000000'",
  responseShape:
    'collection of Microsoft Graph `chatMessage` resources under `value[]`, each with `replyToId` set to the root post: `id`, `messageType`, `createdDateTime`, `deletedDateTime`, `from { user { id, displayName } }`, `body { contentType, content }`, `attachments[]`, `mentions[]`, `reactions[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
