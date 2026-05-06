import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/places/microsoft.graph.room', schema);

const meta: CommandMeta = {
  summary:
    'List bookable meeting rooms in the tenant. Each `room` has `displayName`, `emailAddress`, `capacity`, `building`, `floorNumber`, and `isWheelChairAccessible`. Use the `emailAddress` as a meeting `attendee` for room booking.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/places/microsoft.graph.room',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/place-list',
  options: [],
  example: 'ask-marcel list-rooms',
  responseShape: 'collection of Microsoft Graph `room` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
