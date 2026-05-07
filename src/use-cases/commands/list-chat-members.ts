import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ chatId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/chats/${p.chatId}/members`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the members of a single Microsoft Teams chat.',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '/chats/{chat-id}/members',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list-members',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description:
        'Microsoft Teams chat ID, e.g. `19:abc...@thread.v2`. ' +
        'This CLI does not list chats (the Teams web client token has no `Chat.Read*` scope), so the chat ID has to come from outside. ' +
        'Easiest source from inside this CLI: pick any Teams meeting from `list-calendar-events` and URL-decode the ' +
        '`19%3ameeting_...%40thread.v2` segment of its `onlineMeeting.joinUrl`. ' +
        'Other sources: the Teams desktop / web client (Open in browser → URL contains the chat thread ID), Microsoft Graph Explorer, or Power Automate.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-chat-members --chat-id '19:abc...@thread.v2'",
  responseShape: 'collection of Microsoft Graph `conversationMember` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
