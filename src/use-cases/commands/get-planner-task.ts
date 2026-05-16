import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ plannerTaskId: z.string().min(1) });
const { execute } = buildCommand((p) => `/planner/tasks/${p.plannerTaskId}`, schema);

const meta: CommandMeta = {
  summary: 'Get the metadata of a single Microsoft Planner task (title, assignees, dates, completion).',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/planner/tasks/{planner-task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertask-get',
  options: [
    {
      name: 'planner-task-id',
      key: 'plannerTaskId',
      required: true,
      description:
        "Planner task ID. Returned by `ask-marcel list-planner-tasks` or `list-plan-tasks`. Accepts `--task-id` as a shorter alias (each task command targets exactly one of Planner or To Do, so within this command's flag set there is no ambiguity).",
      aliases: [{ name: 'task-id', key: 'taskId' }],
    },
  ],
  example: "ask-marcel get-planner-task --planner-task-id '01tx7Ic7-USXEwt0lvR1cmgAH8gK'",
  responseShape: 'single Microsoft Graph `plannerTask` resource',
};

export { execute, meta, schema };
