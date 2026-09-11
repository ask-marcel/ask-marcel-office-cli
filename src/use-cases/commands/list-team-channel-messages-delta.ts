import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import { CHANNEL_MESSAGES_TOP_OPTION, withChannelMessagesTopCap } from './channel-message-page.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1), since: isoDateTimeField.optional() });
// Probed live 2026-09-10 on the basic token: the delta route returns ROOT posts
// only, honours `$top` (max 50), `$filter=lastModifiedDateTime gt <instant>`
// (the only filter it takes) and `$expand=replies`; a fresh reply bumps its
// root's `lastModifiedDateTime`, so a thread with new replies comes back.
const sinceFilter = (since: string | undefined): string => (since === undefined ? '' : `?$filter=lastModifiedDateTime%20gt%20${since}`);
const inner = buildPickODataListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/messages/delta${sinceFilter(p.since)}`, baseSchema, ['top', 'expand']);

const execute: Command['execute'] = withChannelMessagesTopCap(async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params)));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "Track what changed in a single channel of a Microsoft Team: the first call returns the root posts (created, edited or replied to, since a fresh reply updates its root) plus `@odata.nextLink` while paging and `@odata.deltaLink` on the final page; feed either to `next-page`, and keep the deltaLink to ask later for only what changed since. Without `--since` the first sync walks the whole channel history; `--since` (an ISO instant or a relative date such as `7d`) bounds it on `lastModifiedDateTime`, the one filter Graph accepts here. Root posts only: pass `--expand replies` to inline each post's replies. `--top` pages up to 50 at a time. Reads through Graph on the basic token; an unknown channel id is named in the error.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages/delta?$filter=lastModifiedDateTime gt {since}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-delta',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    {
      name: 'since',
      key: 'since',
      required: false,
      description: `Lower bound on \`lastModifiedDateTime\` for the first sync (the deltaLink encodes the position for later syncs). Omit it to sync the whole history. ${RELATIVE_DATE_DESCRIPTION}`,
    },
    CHANNEL_MESSAGES_TOP_OPTION,
    {
      name: 'expand',
      key: 'expand',
      required: false,
      description: "OData $expand: pass `replies` to inline each root post's replies (the only navigation Graph expands here). Increases response size.",
    },
  ],
  example: "ask-marcel-office list-team-channel-messages-delta --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --since '7d' --expand replies",
  responseShape:
    'collection of Microsoft Graph `chatMessage` root posts under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page; feed either to `next-page`.',
  pagination: true,
  paginationStrategy: 'deltaLink',
};

export { execute, meta, schema };
