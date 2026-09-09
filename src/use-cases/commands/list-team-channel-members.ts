import { z } from 'zod';
import { buildFilterSelectListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { filterSelectOptions } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
// Same OData surface as `list-team-members`, probed live 2026-09-09: `$select`
// and `$filter` work, `$skip` is a 400 and `$top` mis-pages.
const inner = buildFilterSelectListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/members`, baseSchema);

const execute: Command['execute'] = async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "List the members of a single channel inside a Microsoft Team: the same `conversationMember` entries as `list-team-members` (`displayName`, `email`, `userId`, `roles`), scoped to the channel. On a standard channel this is the whole team; on a private or shared channel it is the channel roster, which `list-team-members` cannot see. `--filter` narrows server-side (owners only with `roles/any(r:r eq 'owner')` on `microsoft.graph.aadUserConversationMember`); only `--filter` and `--select` reach Graph, a large roster continues through the `next:` footer with `next-page`. An unknown channel id is named in the error instead of the bare `1: NotFound` Graph answers.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/members',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/channel-list-members',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    ...filterSelectOptions,
  ],
  example: "ask-marcel-office list-team-channel-members --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2'",
  responseShape:
    'collection of Microsoft Graph `aadUserConversationMember` resources under `value[]`: `id`, `roles[]`, `displayName`, `userId`, `email`, `tenantId`, `visibleHistoryStartDateTime`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
