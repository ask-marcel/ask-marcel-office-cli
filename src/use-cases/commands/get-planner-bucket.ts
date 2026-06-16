import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ plannerBucketId: z.string().min(1) });
const { execute } = buildCommand((p) => `/planner/buckets/${p.plannerBucketId}`, schema);

const meta: CommandMeta = {
  summary: 'Get the metadata of a single Microsoft Planner bucket (column / lane).',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/planner/buckets/{planner-bucket-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-get',
  options: [
    {
      name: 'planner-bucket-id',
      key: 'plannerBucketId',
      required: true,
      description: 'Planner bucket ID. Returned by `ask-marcel list-plan-buckets`.',
      aliases: [
        { name: 'id', key: 'id' },
        { name: 'bucket-id', key: 'bucketId' },
      ],
    },
  ],
  example: "ask-marcel get-planner-bucket --planner-bucket-id 'sFNeQRFu_kqhxpwwAhmA15gAGfoT'",
  responseShape: 'single Microsoft Graph `plannerBucket` resource',
};

export { execute, meta, schema };
