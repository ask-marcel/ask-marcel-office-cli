import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/organization', schema);

const meta: CommandMeta = {
  summary:
    "Return the tenant's organization metadata — display name, country, verified domains, business phones, technical / security notification contacts, assigned Microsoft 365 SKUs / licensing. Useful for confirming which tenant the CLI is signed into and what subscriptions are active.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/organization',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/organization-list',
  options: [],
  example: 'ask-marcel get-organization',
  responseShape: 'collection of one Microsoft Graph `organization` resource under `value[]`',
};

export { execute, meta, schema };
