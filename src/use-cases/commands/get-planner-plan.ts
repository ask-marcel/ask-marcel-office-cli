import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ plannerPlanId: z.string().min(1) });
const { execute } = buildCommand((p) => `/planner/plans/${p.plannerPlanId}`, schema);

const meta: CommandMeta = {
  summary: 'Get the metadata of a single Microsoft Planner plan (title, owner group, container).',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/planner/plans/{planner-plan-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerplan-get',
  options: [
    {
      name: 'planner-plan-id',
      key: 'plannerPlanId',
      required: true,
      description: 'Planner plan ID. Returned in the `planId` field of any task from `ask-marcel list-planner-tasks`.',
      aliases: [
        { name: 'id', key: 'id' },
        { name: 'plan-id', key: 'planId' },
      ],
    },
  ],
  example: "ask-marcel get-planner-plan --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'",
  responseShape: 'single Microsoft Graph `plannerPlan` resource',
};

export { execute, meta, schema };
