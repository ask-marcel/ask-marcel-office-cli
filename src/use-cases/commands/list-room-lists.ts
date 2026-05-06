import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/places/microsoft.graph.roomList', schema);

const meta: CommandMeta = {
  summary:
    'List room lists — usually one per building. Use these to scope a room search by location: a roomList groups the rooms in one office, then `/places/{roomList}/rooms` lists just those rooms.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/places/microsoft.graph.roomList',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/place-list',
  options: [],
  example: 'ask-marcel list-room-lists',
  responseShape: 'collection of Microsoft Graph `roomList` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
