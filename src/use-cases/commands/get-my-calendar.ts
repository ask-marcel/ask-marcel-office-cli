import { z } from 'zod';
import { buildSelectableCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { selectExpandOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({});
const { execute, schema } = buildSelectableCommand(() => '/me/calendar', baseSchema);

const meta: ReadCommandMeta = {
  summary:
    "Return metadata for the signed-in user's *primary* calendar — `id`, `name`, `color`, `owner`, `canShare`, `canViewPrivateItems`, `canEdit`, `defaultOnlineMeetingProvider`. Sibling to `list-calendars` which returns every calendar (incl. shared / subscribed). Use `--select` to fetch only the fields you need.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/calendar',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get-calendar',
  options: [...selectExpandOptions],
  example: "{read-bin} get-my-calendar --select 'id,name'",
  responseShape: 'single Microsoft Graph `calendar` resource',
  effect: 'read',
  scopesRequired: ['Calendars.Read'],
};

export { execute, meta, schema };
