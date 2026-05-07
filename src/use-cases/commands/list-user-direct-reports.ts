import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/users/${p.userId}/directReports`, baseSchema);

const meta: CommandMeta = {
  summary: "List a specific user's direct reports.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/directReports',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-directreports',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: "Azure AD user ID or userPrincipalName (UPN) — typically the user's email address. Use `list-users` to find one.",
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-user-direct-reports --user-id 'alice@contoso.com'",
  responseShape: 'collection of Microsoft Graph `directoryObject` resources (typically `user`) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
