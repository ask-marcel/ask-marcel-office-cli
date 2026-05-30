import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

/**
 * Enumerate every SharePoint site the signed-in user can access via the Microsoft
 * Search index. `search-sharepoint-sites-by-name` uses `GET /sites?search=`, which
 * returns a single capped page (no `@odata.nextLink`). The Search API
 * (`POST /search/query`, entityTypes: ['site']) exposes the FULL security-trimmed
 * index and reports `total` + `moreResultsAvailable`, so this command deep-pages it
 * with `from`/`size` until the index is exhausted (or a page ceiling is hit),
 * deduping site resources by id. Pairs with `list-accessible-drives` (membership /
 * sharing / channels) — the union of the two is the practical delegated maximum.
 */

const PAGE_SIZE = 25;
const MAX_PAGES = 60;

const schema = z.object({ query: z.string().min(1).optional() });

type Hit = { readonly resource?: unknown };
type HitsContainer = { readonly total?: unknown; readonly moreResultsAvailable?: unknown; readonly hits?: ReadonlyArray<Hit> };
type SearchResponse = { readonly value?: ReadonlyArray<{ readonly hitsContainers?: ReadonlyArray<HitsContainer> }> };

const firstContainer = (body: unknown): HitsContainer | undefined => (body as SearchResponse | null)?.value?.[0]?.hitsContainers?.[0];

const siteId = (resource: unknown): string | undefined => {
  if (resource === null || typeof resource !== 'object') return undefined;
  const id = (resource as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryString = parsed.data.query ?? '*';

  const seen = new Set<string>();
  const value: Array<unknown> = [];
  let total = 0;
  let truncated = false;
  for (let page = 0; page < MAX_PAGES; page += 1) {
    const r = await graph.post('/search/query', { requests: [{ entityTypes: ['site'], query: { queryString }, from: page * PAGE_SIZE, size: PAGE_SIZE }] });
    if (!r.ok) {
      if (page === 0) return r;
      truncated = true;
      break;
    }
    const container = firstContainer(r.value);
    if (typeof container?.total === 'number') total = container.total;
    for (const hit of container?.hits ?? []) {
      const id = siteId(hit.resource);
      if (id !== undefined && !seen.has(id)) {
        seen.add(id);
        value.push(hit.resource);
      }
    }
    if (container?.moreResultsAvailable !== true) break;
    if (page === MAX_PAGES - 1) truncated = true;
  }

  return ok({ value, count: value.length, total, ...(truncated ? { truncated: true } : {}) });
};

const meta: CommandMeta = {
  summary:
    'Enumerate EVERY SharePoint site the signed-in user can access via the Microsoft Search index — far more than `search-sharepoint-sites-by-name`, which calls `GET /sites?search=` and returns a single capped page with no continuation. This command deep-pages the Search API (`POST /search/query` with `entityTypes: ["site"]`) using `from`/`size`, following the index\'s own `moreResultsAvailable` flag until exhausted (or the page ceiling of 60×25 = 1500 is reached, signalled by `truncated: true`), and dedupes site resources by id. The index is security-trimmed, so it returns sites you can open even when you are not a member (the gap `list-accessible-drives` cannot fill). Conversely it does NOT return OneDrives, private channel sites, or direct-link-only sites — so the *union of this command and `list-accessible-drives` is the practical maximum reachable on a delegated token* (a truly exhaustive list of every site in the tenant needs admin-only app permissions: `GET /sites/getAllSites`). Optional `--query` narrows the index (default `*` = all accessible sites).',
  category: 'sharepoint',
  graphMethod: 'POST',
  graphPathTemplate: '/search/query',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/search-query',
  options: [
    {
      name: 'query',
      key: 'query',
      required: false,
      description:
        'Optional KQL filter applied to the site index (default `*` = every site you can access). Examples: a name fragment like `budget`, or `contentclass:STS_Site` to restrict to site collections. Free text is matched against site title/url.',
    },
  ],
  example: 'ask-marcel search-all-accessible-sites --output json',
  bodyTemplate:
    "{ requests: [{ entityTypes: ['site'], query: { queryString: '{query}' }, from: <page*25>, size: 25 }] } — `{query}` defaults to `*` (all accessible sites); re-issued per page, advancing `from` by 25 until `moreResultsAvailable` is false",
  responseShape:
    "`{ value: [<Microsoft Graph site resource: { id, name, displayName?, webUrl, … }>], count, total, truncated?: true }`. `value[]` is deduped by site `id` across pages; `count` is how many distinct sites were returned; `total` is the index's own reported match count (compare with `count` to see if the ceiling truncated the sweep); `truncated: true` means paging stopped early (page ceiling hit, or a later page errored) — narrow with `--query` to see the rest.",
};

export { execute, meta, schema };
