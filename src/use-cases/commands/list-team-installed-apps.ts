import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ teamId: z.string().min(1) });
// `$expand=teamsAppDefinition` is hard-pinned because the bare endpoint
// returns only `{ id, consentedPermissionSet }` per app — opaque base64
// IDs with no human-readable signal. The expand surfaces `displayName`,
// `version`, and `distributionMethod` on every entry. Graph rejects the
// other OData passthroughs on this endpoint (`Query option 'Top' is not
// allowed`) so they remain unexposed.
const { execute } = buildCommand((p) => `/teams/${p.teamId}/installedApps?$expand=teamsAppDefinition`, schema);

const meta: CommandMeta = {
  summary:
    "List the Teams apps installed in a team. The CLI hard-pins `$expand=teamsAppDefinition` so every entry includes `displayName`, `version`, and `distributionMethod` (the bare endpoint returns only opaque IDs). Useful for surfacing which integrations are wired into a given team. Graph rejects user-supplied OData query parameters on this endpoint (`Query option 'Top' is not allowed`) — so the standard OData flags are intentionally NOT exposed here. The response itself is still server-paginated via `@odata.nextLink` when the team has many installed apps; chain with `next-page` to walk subsequent pages.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/installedApps?$expand=teamsAppDefinition',
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
  responseShape:
    'collection of Microsoft Graph `teamsAppInstallation` resources under `value[]`, each with an inline `teamsAppDefinition` (`displayName`, `version`, `distributionMethod`)',
  pagination: true,
};

export { execute, meta, schema };
