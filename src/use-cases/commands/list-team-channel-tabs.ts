import { z } from 'zod';
import { buildFilterSelectListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { filterSelectOptions } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
// `$expand=teamsApp` is hard-pinned, as `list-team-installed-apps` pins its
// definition: the bare entry carries only `id`, `displayName`, `webUrl` and a
// `configuration` blob, and the app behind a tab (Planner, a website, a
// SharePoint page, a OneNote section) is what a reader needs to know what the
// tab is. Probed live 2026-09-09: `$select` and `$filter` work alongside the
// expand; `$top` is a 400 (`Query option 'Top' is not allowed`), and a channel
// has few tabs, so there is no paging to advertise.
const inner = buildFilterSelectListCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/tabs?$expand=teamsApp`, baseSchema);

const execute: Command['execute'] = async (graph, params) => rewriteChannelScopedError(await inner.execute(graph, params), channelScopeOf(params));
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the tabs pinned to a single channel of a Microsoft Team, each with its `teamsApp` expanded (the CLI hard-pins `$expand=teamsApp`) so a tab reads as "Planner", "Website" or "SharePoint" and not only as its `displayName`. `webUrl` opens the tab in Teams; `configuration` carries the app\'s own settings (a Planner `planId`, a website `websiteUrl`, a Files folder `contentUrl`). `--filter` narrows server-side (`displayName eq \'Files\'`) and `--select` trims the projection; Graph rejects `$top` here and a channel has few tabs, so there is no paging. An unknown channel id is named in the error instead of the Skype backend failure Graph answers.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/tabs?$expand=teamsApp',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/channel-list-tabs',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    ...filterSelectOptions,
  ],
  example: "ask-marcel-office list-team-channel-tabs --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2'",
  responseShape:
    'collection of Microsoft Graph `teamsTab` resources under `value[]`: `id`, `displayName`, `webUrl`, `configuration { entityId, contentUrl, websiteUrl, removeUrl }` and an inline `teamsApp { id, externalId, displayName, distributionMethod }`',
};

export { execute, meta, schema };
