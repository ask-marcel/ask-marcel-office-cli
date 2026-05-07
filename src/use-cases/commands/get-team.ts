import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/teams/${p.teamId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get the metadata of a single Microsoft Team (display name, settings, member-settings, owner group). Pass `--select displayName,description,visibility` to slim the response.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/team-get',
  options: [{ name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel list-joined-teams`.' }, ...selectExpandOptions],
  example: "ask-marcel get-team --team-id 'abc-1234-...' --select displayName,description,visibility",
  responseShape: 'single Microsoft Graph `team` resource (or projection of the requested `--select` fields)',
};

export { execute, meta, schema };
