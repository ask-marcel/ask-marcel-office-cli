import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/manager', schema);

const meta: CommandMeta = {
  summary:
    "Return the signed-in user's manager (a single `user` resource). Returns 404 `Request_ResourceNotFound` if no manager is set in the directory — that is data-empty, not a permission failure.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/manager',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-manager',
  options: [],
  example: 'ask-marcel get-my-manager',
  responseShape: 'single Microsoft Graph `user` resource',
};

export { execute, meta, schema };
