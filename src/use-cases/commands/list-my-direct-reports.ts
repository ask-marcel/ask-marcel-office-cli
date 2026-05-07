import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/directReports', baseSchema);

const meta: CommandMeta = {
  summary: "List the signed-in user's direct reports (employees who report to them in the directory).",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/directReports',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-directreports',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-my-direct-reports',
  responseShape: 'collection of Microsoft Graph `directoryObject` resources (typically `user`) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
