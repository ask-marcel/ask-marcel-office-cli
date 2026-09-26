import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ plannerPlanId: z.string().min(1) });
const { execute } = buildCommand((p) => `/planner/plans/${p.plannerPlanId}/details`, schema);

const meta: CommandMeta = {
  summary:
    "Get the details of a Microsoft Planner plan: the names of its labels (`categoryDescriptions`, what `category1` to `category25` in a task's `appliedCategories` stand for) and who the plan is shared with.",
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/planner/plans/{planner-plan-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerplandetails-get',
  options: [
    {
      name: 'planner-plan-id',
      key: 'plannerPlanId',
      required: true,
      description: 'Planner plan ID. Returned by `ask-marcel-office list-planner-plans` or `list-group-planner-plans`, and in the `planId` field of any task.',
    },
  ],
  example: "ask-marcel-office get-planner-plan-details --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'",
  responseShape: 'single Microsoft Graph `plannerPlanDetails` resource',
};

export { execute, meta, schema };
