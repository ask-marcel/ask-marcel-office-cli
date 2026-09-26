import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectOnlyOptions } from './odata-query.ts';

// A group's plans are the group's: belonging to the group is what grants them,
// and `/me/planner/plans` need not list them. Verified live 2026-09-12 by a
// consumer: `list-planner-plans` answered no plans while the signed-in user's
// four Microsoft 365 groups held three between them, all reachable here. Only
// `--select` is advertised, as on `list-planner-plans`, so the manifest promises
// only what the plans collections are known to honour.
const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildPickODataListCommand((p) => `/groups/${p.groupId}/planner/plans`, baseSchema, ['select']);

const meta: CommandMeta = {
  summary:
    "List the Microsoft Planner plans a Microsoft 365 group owns. A group's plans are granted by membership and need not appear in `list-planner-plans`, which lists the plans shared with the signed-in user: to find every plan a user can read, list their groups with `list-my-memberships` and ask each here. Only `--select` is advertised; slice and sort client-side.",
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/planner/plans',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID of a Microsoft 365 group. Returned by `ask-marcel-office list-my-memberships` or `list-groups`.',
    },
    ...selectOnlyOptions,
  ],
  example: "ask-marcel-office list-group-planner-plans --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `plannerPlan` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
