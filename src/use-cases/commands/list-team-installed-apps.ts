import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ teamId: z.string().min(1) });
const { execute } = buildCommand((p) => `/teams/${p.teamId}/installedApps`, schema);

const meta: CommandMeta = {
  summary:
    'List the Teams apps installed in a team (incl. teamsAppDefinition `displayName`, `version`, `distributionMethod`). Useful for surfacing which integrations are wired into a given team.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/installedApps',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/team-list-installedapps',
  options: [
    {
      name: 'team-id',
      key: 'teamId',
      required: true,
      description: 'Microsoft Teams team ID.',
    },
  ],
  example: "ask-marcel list-team-installed-apps --team-id 'tm1'",
  responseShape: 'collection of Microsoft Graph `teamsAppInstallation` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
