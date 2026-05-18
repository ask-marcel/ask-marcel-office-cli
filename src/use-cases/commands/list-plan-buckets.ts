import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectOnlyOptions } from './odata-query.ts';

// Same Graph behavior as `/me/planner/plans`: `$select` works, everything
// else (`$top`, `$skip`, `$filter`, `$orderby`) is silently dropped.
const baseSchema = z.object({ plannerPlanId: z.string().min(1) });
const { execute, schema } = buildPickODataListCommand((p) => `/planner/plans/${p.plannerPlanId}/buckets`, baseSchema, ['select']);

const meta: CommandMeta = {
  summary:
    'List the buckets (columns / lanes) of a Microsoft Planner plan. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — slice / sort client-side.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/planner/plans/{planner-plan-id}/buckets',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerplan-list-buckets',
  options: [
    {
      name: 'planner-plan-id',
      key: 'plannerPlanId',
      required: true,
      description: 'Planner plan ID. Returned in the `planId` field of any task from `ask-marcel list-planner-tasks`.',
      aliases: [{ name: 'plan-id', key: 'planId' }],
    },
    ...selectOnlyOptions,
  ],
  example: "ask-marcel list-plan-buckets --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'",
  responseShape: 'collection of Microsoft Graph `plannerBucket` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
