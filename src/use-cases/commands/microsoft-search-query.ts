import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const ALL_ENTITY_TYPES = ['driveItem', 'listItem', 'site', 'message', 'event', 'person'] as const;
type EntityType = (typeof ALL_ENTITY_TYPES)[number];
const PAGE_SIZE = 25;

const schema = z.object({ query: z.string().min(1) });

type SearchHitsContainer = { readonly searchTerms?: ReadonlyArray<string>; readonly hitsContainers?: ReadonlyArray<unknown> };
type SearchResponse = { readonly value?: ReadonlyArray<SearchHitsContainer> };

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryString = parsed.data.query;

  const calls = ALL_ENTITY_TYPES.map((entityType) => graph.post('/search/query', { requests: [{ entityTypes: [entityType], query: { queryString }, size: PAGE_SIZE }] }));
  const results = await Promise.all(calls);

  const merged: SearchHitsContainer[] = [];
  const partialErrors: { readonly entityType: EntityType; readonly error: GraphError }[] = [];
  for (let i = 0; i < results.length; i += 1) {
    const r = results[i];
    const entityType = ALL_ENTITY_TYPES[i];
    if (entityType === undefined) continue;
    if (r === undefined) continue;
    if (r.ok) {
      const innerValue = (r.value as SearchResponse).value;
      if (Array.isArray(innerValue)) merged.push(...innerValue);
    } else {
      partialErrors.push({ entityType, error: r.error });
    }
  }

  if (merged.length === 0 && partialErrors.length > 0) return err(partialErrors[0]?.error ?? { type: 'api_error', status: 500, message: 'all entity-type sub-requests failed' });

  return ok({ value: merged, ...(partialErrors.length > 0 ? { partialErrors } : {}) });
};

const meta: CommandMeta = {
  summary:
    "Run a federated KQL search across the signed-in user's mail, files, list items, sites, calendar events, and people. Microsoft Graph v1.0 rejects multi-entity search bodies on most tenants (`Multiple entity search is not supported in v1.0`), so this command issues SIX parallel POSTs — one per entityType — and merges the per-entity `searchHits` containers into a single `value[]`. Each container is identifiable by the resource type inside `hits[].resource`. If a sub-request fails (e.g. tenant lacks the scope for one entity), the others still return; failures show up in `partialErrors[]`. Page size is fixed at 25 per sub-request and `top` is NOT exposed (Graph rejects $top in /search/query bodies). `chatMessage` is excluded since `Chat.Read*` is unavailable.",
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
    "{ requests: [{ entityTypes: ['<one-of-driveItem-listItem-site-message-event-person>'], query: { queryString: '{query}' }, size: 25 }] } — sent six times in parallel, one per entityType",
  responseShape:
    'merged Microsoft Graph `searchResponse` envelope: `{ value: [{ searchTerms, hitsContainers: [{ total, hits: [{ hitId, rank, summary, resource }] }] }, …], partialErrors?: [{ entityType, error }] }`. value[] holds one container per entityType that succeeded; partialErrors[] (only present when at least one sub-request failed) lists which entityTypes returned errors.',
};

export { execute, meta, schema };
