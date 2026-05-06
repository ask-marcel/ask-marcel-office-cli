import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ userId: z.string().min(1) });
const { execute } = buildCommand((p) => `/users/${p.userId}/manager`, schema);

const meta: CommandMeta = {
  summary: "Return a specific user's manager (a single `user` resource). 404 if no manager is set in the directory.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/manager',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-manager',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN. Use `list-users` to find one.',
    },
  ],
  example: "ask-marcel get-user-manager --user-id 'alice@contoso.com'",
  responseShape: 'single Microsoft Graph `user` resource',
};

export { execute, meta, schema };
