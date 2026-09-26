import { z } from 'zod';
import { buildNoSkipListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { noSkipOptions, noSkipShape } from './odata-query.ts';
import { recentItemPath, WITH_ITEM_OPTION, withItemEnrichment } from './with-item.ts';

const baseSchema = z.object({}).strict();
const inner = buildNoSkipListCommand(() => '/me/drive/recent', baseSchema);
const schema = z.object({ ...baseSchema.shape, ...noSkipShape, withItem: z.enum(['true', 'false']).optional() });
const execute = withItemEnrichment(schema, inner.execute, recentItemPath);

const meta: CommandMeta = {
  summary:
    'List the signed-in user\'s most recently used / opened OneDrive and SharePoint files, ranked by Microsoft\'s recency signal. The strongest single answer to "what is this user working on right now?". Note: Graph\'s recent-files feed is signal-driven and can lag the underlying drive by 24-48 hours — `lastModifiedDateTime` here may be older than the file\'s true mtime. For "what is the actual latest version?" call `list-drive-item-versions` on a specific item.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/recent',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-recent',
  options: [...noSkipOptions, WITH_ITEM_OPTION],
  example: 'ask-marcel-office list-recent-files',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
