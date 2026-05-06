import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const DEFAULT_ENTITY_TYPES = ['driveItem', 'listItem', 'site', 'message', 'event', 'person'] as const;

const schema = z.object({ query: z.string().min(1) });

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const body = {
    requests: [
      {
        entityTypes: DEFAULT_ENTITY_TYPES,
        query: { queryString: parsed.data.query },
        size: 25,
      },
    ],
  };
  return graph.post('/search/query', body);
};

const meta: CommandMeta = {
  summary:
    "Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people in one round trip — Microsoft's unified search API, the same engine that powers the Microsoft 365 search box. Each `searchHits[]` entry has `_score`, `summary`, and a typed `resource`. Page size is fixed at 25 in this command; for larger pages or different entity types call the Graph endpoint directly via the library API. `chatMessage` is intentionally omitted from the default entity set since `Chat.Read*` is unavailable.",
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
  bodyTemplate: "{ requests: [{ entityTypes: ['driveItem','listItem','site','message','event','person'], query: { queryString: '{query}' }, size: 25 }] }",
  responseShape: 'Microsoft Graph `searchResponse` envelope: `{ value: [{ searchTerms, hitsContainers: [{ total, hits: [{ hitId, rank, summary, resource }] }] }] }`',
};

export { execute, meta, schema };
