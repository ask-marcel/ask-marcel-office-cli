import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Deep-page `POST /search/query` over files (`entityTypes: ['driveItem']`), 200
 * hits a page (Microsoft's recommended driveItem page size), following the
 * index's own `moreResultsAvailable` flag until it is exhausted or `maxPages`
 * pages have been read. Hits are deduped by `hitId` (driveItem ids are only
 * unique within a drive, so `hitId` is the cross-drive key; falls back to
 * `resource.id`). A failure on the first page is the caller's error; a failure
 * later, or the page ceiling, ends the sweep with `truncated: true`.
 */

const PAGE_SIZE = 200;

type Hit = { readonly hitId?: unknown; readonly resource?: unknown };
type HitsContainer = { readonly moreResultsAvailable?: unknown; readonly hits?: ReadonlyArray<Hit> };
type SearchResponse = { readonly value?: ReadonlyArray<{ readonly hitsContainers?: ReadonlyArray<HitsContainer> }> };
type DriveItemSweep = { readonly value: ReadonlyArray<unknown>; readonly truncated: boolean };

const firstContainer = (body: unknown): HitsContainer | undefined => (body as SearchResponse | null)?.value?.[0]?.hitsContainers?.[0];

const dedupKey = (hit: Hit): string | undefined => {
  if (typeof hit.hitId === 'string') return hit.hitId;
  const resource = hit.resource;
  if (resource === null || typeof resource !== 'object') return undefined;
  const id = (resource as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const searchDriveItems = async (graph: GraphClient, queryString: string, maxPages: number): Promise<Result<DriveItemSweep, GraphError>> => {
  const seen = new Set<string>();
  const value: Array<unknown> = [];
  let truncated = false;
  for (let page = 0; page < maxPages; page += 1) {
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
    if (page === maxPages - 1) truncated = true;
  }
  return ok({ value, truncated });
};

export { searchDriveItems };
