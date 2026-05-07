import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const FILE_MAIL_EVENT_TYPES = ['driveItem', 'listItem', 'site', 'message', 'event'] as const;
const PEOPLE_TYPES = ['person'] as const;
const PAGE_SIZE = 25;

const schema = z.object({ query: z.string().min(1) });

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryString = parsed.data.query;
  const body = {
    requests: [
      { entityTypes: FILE_MAIL_EVENT_TYPES, query: { queryString }, size: PAGE_SIZE },
      { entityTypes: PEOPLE_TYPES, query: { queryString }, size: PAGE_SIZE },
    ],
  };
  return graph.post('/search/query', body);
};

const meta: CommandMeta = {
  summary:
    "Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people. Microsoft Graph rejects mixing `person` with file/mail/event types in a single request, so this command sends two `requests[]` entries in one search body — one for files/mail/events, one for people — and returns Graph's response unchanged. `value[0]` holds files/mail/events hits; `value[1]` holds people hits. Each `searchHits[]` entry has `_score`, `summary`, and a typed `resource`. Page size is fixed at 25 per sub-request. `chatMessage` is intentionally omitted from the entity set since `Chat.Read*` is unavailable.",
  category: 'meta',
  graphMethod: 'POST',
  graphPathTemplate: '/search/query',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/search-query',
  options: [
    {
      name: 'query',
      key: 'query',
      required: true,
      description:
        'KQL query string. Supports field operators where indexed by the corpus (e.g. `from:alice`, `subject:"q3 budget"`, `filetype:xlsx`). Free-text works everywhere.',
    },
  ],
  example: "ask-marcel microsoft-search-query --query 'q3 budget'",
  bodyTemplate:
    "{ requests: [{ entityTypes: ['driveItem','listItem','site','message','event'], query: { queryString: '{query}' }, size: 25 }, { entityTypes: ['person'], query: { queryString: '{query}' }, size: 25 }] }",
  responseShape:
    'Microsoft Graph `searchResponse` envelope: `{ value: [{ searchTerms, hitsContainers: [{ total, hits: [{ hitId, rank, summary, resource }] }] }, …] }`. `value[0]` = files/mail/events, `value[1]` = people.',
};

export { execute, meta, schema };
