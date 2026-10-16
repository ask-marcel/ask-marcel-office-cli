import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/places/microsoft.graph.room', baseSchema);

const meta: ReadCommandMeta = {
  summary:
    'List bookable meeting rooms in the tenant. Each `room` has `displayName`, `emailAddress`, `capacity`, `building`, `floorNumber`, and `isWheelChairAccessible`. Use the `emailAddress` as a meeting `attendee` for room booking. Pass `--top 5` to limit the response — large tenants return tens of KB by default.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/places/microsoft.graph.room',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/place-list',
  options: [...odataQueryOptions],
  example: '{read-bin} list-rooms',
  responseShape: 'collection of Microsoft Graph `room` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Place.Read.All'],
};

export { execute, meta, schema };
