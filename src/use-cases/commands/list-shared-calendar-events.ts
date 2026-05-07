import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/users/${p.userId}/calendar/events`, baseSchema);

const meta: CommandMeta = {
  summary: "List events from another user's primary calendar (shared / delegated access). 403 without `Calendars.Read.Shared`.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/calendar/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-events',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN whose calendar to read. Requires `Calendars.Read.Shared` access (granted by the calendar owner).',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-shared-calendar-events --user-id 'colleague@contoso.com'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
