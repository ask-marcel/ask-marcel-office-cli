import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildNoSkipListCommand(() => '/groups', baseSchema);

const meta: CommandMeta = {
  summary: 'List Microsoft 365 groups, security groups, and distribution groups in the tenant directory. Use `--top` and `next-page` to paginate over very large directories.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/groups',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list',
  options: [...noSkipOptions],
  example: 'ask-marcel list-groups',
  responseShape: 'collection of Microsoft Graph `group` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
