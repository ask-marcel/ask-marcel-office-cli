import { z } from 'zod';
import { buildNoSkipListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { noSkipOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildNoSkipListCommand(() => '/groups', baseSchema);

const meta: ReadCommandMeta = {
  summary: 'List Microsoft 365 groups, security groups, and distribution groups in the tenant directory. Use `--top` and `next-page` to paginate over very large directories.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/groups',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list',
  options: [...noSkipOptions],
  example: '{read-bin} list-groups',
  responseShape: 'collection of Microsoft Graph `group` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
  effect: 'read',
  scopesRequired: ['GroupMember.Read.All'],
};

export { execute, meta, schema };
