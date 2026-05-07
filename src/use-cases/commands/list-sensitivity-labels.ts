import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/informationProtection/sensitivityLabels', baseSchema);

const meta: CommandMeta = {
  summary:
    'List the Microsoft Information Protection sensitivity labels available to the signed-in user — the labels Outlook / Word / SharePoint surfaces in the "Sensitivity" picker (e.g. Public / Internal / Confidential / Highly Confidential). Each label has `id`, `displayName`, `priority`, `isAppliable`, `tooltip`.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/informationProtection/sensitivityLabels',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/security-informationprotection-list-sensitivitylabels',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-sensitivity-labels',
  responseShape: 'collection of Microsoft Graph `sensitivityLabel` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
