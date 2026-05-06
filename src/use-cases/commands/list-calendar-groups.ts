import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/calendarGroups', schema);

const meta: CommandMeta = {
  summary:
    'List the signed-in user\'s calendar groups — Outlook\'s organizational layer above individual calendars (e.g. "My Calendars", "Other Calendars", "Birthdays"). Use the returned `id` with `list-calendar-group-calendars` to drill in.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendarGroups',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-calendargroups',
  options: [],
  example: 'ask-marcel list-calendar-groups',
  responseShape: 'collection of Microsoft Graph `calendarGroup` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
