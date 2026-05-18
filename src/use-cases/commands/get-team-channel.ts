import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildSelectableCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
const inner = buildSelectableCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}`, baseSchema);

// Graph surfaces a stripped `1: NotFound` for a missing channel-id (the
// `1:` prefix is the Teams thread-id segment, unhelpfully echoed). The
// sibling `get-team` returns a clear `BadRequest: teamId needs to be a
// valid GUID.` Rewrite for parity.
const execute: Command['execute'] = async (graph, params) => {
  const result = await inner.execute(graph, params);
  if (result.ok) return result;
  if (result.error.type === 'api_error' && /^1:\s*NotFound/i.test(result.error.message)) {
    const channelId = typeof params['channelId'] === 'string' ? params['channelId'] : '<unknown>';
    return err({
      type: 'api_error',
      status: result.error.status,
      message: `NotFound: Microsoft Teams channel not found (channel-id: "${channelId}"). Verify it exists in this team via \`ask-marcel list-team-channels --team-id <team-id>\`.`,
      code: 'cli_rewrite_channel_not_found',
    });
  }
  return result;
};
const { schema } = inner;

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
