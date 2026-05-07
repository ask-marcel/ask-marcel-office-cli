import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildSelectableCommand(() => '/me', baseSchema);

const meta: CommandMeta = {
  summary:
    'Return the signed-in user’s Microsoft Graph profile (id, displayName, mail, jobTitle, etc.). Pass `--select id,displayName,mail` to slim the payload to just the fields you need.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [...selectExpandOptions],
  example: 'ask-marcel get-current-user --select id,displayName,mail',
  responseShape: 'single Microsoft Graph `user` resource (or projection of the requested `--select` fields)',
};

export { execute, meta, schema };
