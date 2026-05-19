import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({});
const { execute, schema } = buildSelectableCommand(() => '/organization', baseSchema);

const meta: CommandMeta = {
  summary:
    "Return the tenant's organization metadata — display name, country, verified domains, business phones, technical / security notification contacts, assigned Microsoft 365 SKUs / licensing. Graph wraps the single organization resource under `value[]` (audit v1.0.0 §D7 — even though only one tenant exists, the endpoint returns a collection). The full resource is ~57 KB; use `--select` to slim it (e.g. `--select id,displayName,verifiedDomains`).",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/organization',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/organization-list',
  options: [...selectExpandOptions],
  example: "ask-marcel get-organization --select 'id,displayName,verifiedDomains'",
  responseShape: 'collection of one Microsoft Graph `organization` resource under `value[]`',
};

export { execute, meta, schema };
