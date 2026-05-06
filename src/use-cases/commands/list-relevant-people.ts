import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/people', schema);

const meta: CommandMeta = {
  summary:
    "List people relevant to the signed-in user — colleagues they email and meet with most. Microsoft's relevance ranking, not the full directory. Returns `displayName`, `emailAddresses`, `jobTitle`, `companyName`, etc.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/people',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-people',
  options: [],
  example: 'ask-marcel list-relevant-people',
  responseShape: 'collection of Microsoft Graph `person` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
