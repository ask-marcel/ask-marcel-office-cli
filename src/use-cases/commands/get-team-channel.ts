import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get the metadata of a single channel inside a Microsoft Team. Use `--select` to slim the response (e.g. `--select id,displayName,webUrl`) — sibling to `get-team` and `get-team-primary-channel` which both expose the same flag.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/channel-get',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID. Returned by `ask-marcel list-team-channels`.' },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-team-channel --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --select 'id,displayName'",
  responseShape: 'single Microsoft Graph `channel` resource',
};

export { execute, meta, schema };
