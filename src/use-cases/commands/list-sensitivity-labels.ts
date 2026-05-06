import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/informationProtection/sensitivityLabels', schema);

const meta: CommandMeta = {
  summary:
    'List the Microsoft Information Protection sensitivity labels available to the signed-in user — the labels Outlook / Word / SharePoint surfaces in the "Sensitivity" picker (e.g. Public / Internal / Confidential / Highly Confidential). Each label has `id`, `displayName`, `priority`, `isAppliable`, `tooltip`.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/informationProtection/sensitivityLabels',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/security-informationprotection-list-sensitivitylabels',
  options: [],
  example: 'ask-marcel list-sensitivity-labels',
  responseShape: 'collection of Microsoft Graph `sensitivityLabel` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
