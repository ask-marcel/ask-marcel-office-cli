import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/inferenceClassification/overrides', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Focused Inbox classification overrides — sender addresses they've manually moved to Focused or Other, which override Microsoft's automatic classifier.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/inferenceClassification/overrides',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/inferenceclassification-list-overrides',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-focused-inbox-overrides',
  responseShape: 'collection of Microsoft Graph `inferenceClassificationOverride` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
