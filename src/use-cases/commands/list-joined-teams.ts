import { z } from 'zod';
import { buildCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/joinedTeams', schema);

const meta: ReadCommandMeta = {
  summary:
    "List the Microsoft Teams the signed-in user is a member of. Note: this endpoint does NOT accept the standard OData query parameters — Graph rejects `$top`/`$select`/`$filter`/etc. on `/me/joinedTeams` with `Query option 'X' is not allowed`. The CLI omits the OData passthrough on this command for that reason; pass post-processing through `jq` instead if you need to slice the response.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/me/joinedTeams',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams',
  options: [],
  example: '{read-bin} list-joined-teams',
  responseShape: 'collection of Microsoft Graph `team` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Team.ReadBasic.All'],
};

export { execute, meta, schema };
