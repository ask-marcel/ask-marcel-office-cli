import { z } from 'zod';
import { buildNoSkipListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { noSkipOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildNoSkipListCommand(() => '/me/drive/following', baseSchema);

const meta: ReadCommandMeta = {
  summary:
    'List driveItems the signed-in user has explicitly followed (the OneDrive star). A small, hand-curated set of frequently-revisited files, distinct from the algorithmic `list-recent-files` and `list-recently-used-insights`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/following',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list-following',
  options: [...noSkipOptions],
  example: '{read-bin} list-followed-drive-items',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
  effect: 'read',
  scopesRequired: ['Files.Read'],
};

export { execute, meta, schema };
