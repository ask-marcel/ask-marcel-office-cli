import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

/**
 * Search EVERY file the signed-in user can access — their personal OneDrive, files
 * shared with them, and every SharePoint and Teams (channel) document library they
 * can open — via the Microsoft Search index.
 * `search-my-documents` (`GET /me/drive/search`) sees only the personal OneDrive;
 * `search-onedrive-files` needs a specific `--drive-id`; `microsoft-search-query`
 * reaches everything but bundles mail/calendar/people/sites into the same result.
 * This command is the files-only, all-drives slice: it issues `POST /search/query`
 * with `entityTypes: ['driveItem']` and deep-pages `from`/`size` (200 per page —
 * Microsoft's recommended driveItem page size) following the index's own
 * `moreResultsAvailable` flag until exhausted, or the ceiling of 25×200 = 5000 is
 * reached (`truncated: true`). Hits are deduped by `hitId` (driveItem ids are only
 * unique within a drive, so `hitId` is the cross-drive key; falls back to
 * `resource.id`). The index is security-trimmed, so it returns files across sites
 * the user can open even without membership.
 */

const PAGE_SIZE = 200; // Microsoft's recommended driveItem page size (max is 1000) — balances latency vs round-trips.
const MAX_PAGES = 25; // runaway guard: 25 × 200 = 5000 files, then `truncated: true`. Narrow with --query to see the rest.

const schema = z.object({ query: z.string().min(1) });

type Hit = { readonly hitId?: unknown; readonly resource?: unknown };
type HitsContainer = { readonly moreResultsAvailable?: unknown; readonly hits?: ReadonlyArray<Hit> };
type SearchResponse = { readonly value?: ReadonlyArray<{ readonly hitsContainers?: ReadonlyArray<HitsContainer> }> };

const firstContainer = (body: unknown): HitsContainer | undefined => (body as SearchResponse | null)?.value?.[0]?.hitsContainers?.[0];

const dedupKey = (hit: Hit): string | undefined => {
  if (typeof hit.hitId === 'string') return hit.hitId;
  const resource = hit.resource;
  if (resource === null || typeof resource !== 'object') return undefined;
  const id = (resource as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryString = parsed.data.query;

  const seen = new Set<string>();
  const value: Array<unknown> = [];
  let truncated = false;
  for (let page = 0; page < MAX_PAGES; page += 1) {
    const r = await graph.post('/search/query', { requests: [{ entityTypes: ['driveItem'], query: { queryString }, from: page * PAGE_SIZE, size: PAGE_SIZE }] });
    if (!r.ok) {
      if (page === 0) return r;
      truncated = true;
      break;
    }
    const container = firstContainer(r.value);
    for (const hit of container?.hits ?? []) {
      const key = dedupKey(hit);
      if (key !== undefined && !seen.has(key)) {
        seen.add(key);
        value.push(hit.resource);
      }
    }
    if (container?.moreResultsAvailable !== true) break;
    if (page === MAX_PAGES - 1) truncated = true;
  }

  return ok({ value, count: value.length, ...(truncated ? { truncated: true } : {}) });
};

const meta: CommandMeta = {
  summary:
    'Search EVERY file the signed-in user can access — their personal OneDrive, files shared with them, and every SharePoint and Teams (channel) document library they can open — for a free-text query. Unlike `search-my-documents` (personal OneDrive only) or `search-onedrive-files` (one drive by id), this reaches across all accessible drives via the security-trimmed Microsoft Search index; unlike `microsoft-search-query` it returns FILES ONLY (`entityTypes: ["driveItem"]`), not mail/calendar/people/sites. It deep-pages `POST /search/query` with `from`/`size` (200 per page) following the index\'s `moreResultsAvailable` flag until exhausted, or the ceiling of 25×200 = 5000 files is reached (`truncated: true` — narrow with `--query` to see the rest). Hits are deduped by `hitId`. Each returned `driveItem` carries `id` + `parentReference.driveId`, the pair `download-drive-item-content` / `download-drive-item-as-markdown` need to open it.',
  category: 'drive',
  graphMethod: 'POST',
  graphPathTemplate: '/search/query',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/search-query',
  scopesRequired: ['Files.Read', 'Sites.Read.All'],
  options: [
    {
      name: 'query',
      key: 'query',
      required: true,
      description:
        'KQL / free-text search query. Matches filename, content, and metadata across every drive the user can access. Add `filetype:` to narrow (e.g. `q1 budget filetype:xlsx`), or field operators like `filename:`. Free text works everywhere.',
    },
  ],
  example: "ask-marcel-office search-all-files --query 'q1 budget filetype:xlsx'",
  bodyTemplate:
    "{ requests: [{ entityTypes: ['driveItem'], query: { queryString: '{query}' }, from: <page*200>, size: 200 }] } — re-issued per page, advancing `from` by 200 until `moreResultsAvailable` is false or the 25-page ceiling is hit",
  responseShape:
    '`{ value: [<Microsoft Graph driveItem resource: { id, name, webUrl, parentReference: { driveId }, size, … }>], count, truncated?: true }`. `value[]` is deduped by `hitId` across pages; `count` is the number of distinct files returned. `truncated: true` means paging stopped early (page ceiling hit, or a later page errored) — narrow with `--query` to see the rest; its absence means the sweep ran to completion.',
};

export { execute, meta, schema };
