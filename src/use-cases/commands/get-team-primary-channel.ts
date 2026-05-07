import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/teams/${p.teamId}/primaryChannel`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Return the team's primary (General) channel directly without having to list-then-pick. The returned `channel` has `id`, `displayName`, `webUrl`, `email` — feed `id` into `list-team-channels` siblings or `get-channel-files-folder`.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/primaryChannel',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/team-get-primarychannel',
  options: [
    {
      name: 'team-id',
      key: 'teamId',
      required: true,
      description: 'Microsoft Teams team ID. Returned by `list-joined-teams`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-team-primary-channel --team-id 'tm1'",
  responseShape: 'single Microsoft Graph `channel` resource',
};

export { execute, meta, schema };
