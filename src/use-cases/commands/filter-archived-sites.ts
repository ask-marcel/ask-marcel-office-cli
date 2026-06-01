import { errorIndicatesArchived, isArchivedSite } from '../../domain/utilities/archive-status.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Drop archived sites from a list of Microsoft Graph `site` resources by probing
 * each one's metadata (`GET /sites/{id}?$select=id,webUrl,siteCollection`). A site
 * is excluded when Graph reports it archived (`archiveStatus`) or when the probe
 * fails with an archive/locked signal (423 `resourceLocked`). A probe that fails
 * for any other reason keeps the site (we never lose an accessible site to a
 * transient error) and is counted in `probeErrors`. Probing is capped and
 * concurrency-bounded to avoid 429 throttling on large tenants; sites past the cap
 * are kept unprobed and `probeTruncated` is set.
 */

const PROBE_PATH = '?$select=id,webUrl,siteCollection';
const ARCHIVE_PROBE_MAX = 250;
const PROBE_CHUNK = 15;

type FilterOptions = { readonly probeMax?: number; readonly chunkSize?: number };
type FilterResult = { readonly value: ReadonlyArray<unknown>; readonly archivedExcluded: number; readonly probeErrors: number; readonly probeTruncated: boolean };
type Verdict = 'keep' | 'drop' | 'error';

const siteId = (resource: unknown): string | undefined => {
  if (resource === null || typeof resource !== 'object') return undefined;
  const id = (resource as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const errorFields = (e: GraphError): { status?: number; code?: string; message?: string } => ({
  status: e.type === 'api_error' ? e.status : undefined,
  code: e.code,
  message: e.message,
});

const probe = async (graph: GraphClient, resource: unknown): Promise<Verdict> => {
  const id = siteId(resource);
  if (id === undefined) return 'keep';
  const r = await graph.get(`/sites/${id}${PROBE_PATH}`);
  if (r.ok) return isArchivedSite(r.value) ? 'drop' : 'keep';
  return errorIndicatesArchived(errorFields(r.error)) ? 'drop' : 'error';
};

const verdictsFor = async (graph: GraphClient, sites: ReadonlyArray<unknown>, probeMax: number, chunkSize: number): Promise<ReadonlyArray<Verdict>> => {
  const out: Array<Verdict> = [];
  for (let start = 0; start < sites.length; start += chunkSize) {
    const chunk = sites.slice(start, start + chunkSize);
    const done = await Promise.all(chunk.map((site, j) => (start + j < probeMax ? probe(graph, site) : Promise.resolve<Verdict>('keep'))));
    out.push(...done);
  }
  return out;
};

const filterOutArchivedSites = async (graph: GraphClient, sites: ReadonlyArray<unknown>, options?: FilterOptions): Promise<FilterResult> => {
  const probeMax = options?.probeMax ?? ARCHIVE_PROBE_MAX;
  const chunkSize = options?.chunkSize ?? PROBE_CHUNK;
  const verdicts = await verdictsFor(graph, sites, probeMax, chunkSize);
  const value: Array<unknown> = [];
  let archivedExcluded = 0;
  let probeErrors = 0;
  sites.forEach((site, i) => {
    if (verdicts[i] === 'drop') {
      archivedExcluded += 1;
      return;
    }
    if (verdicts[i] === 'error') probeErrors += 1;
    value.push(site);
  });
  return { value, archivedExcluded, probeErrors, probeTruncated: sites.length > probeMax };
};

export { ARCHIVE_PROBE_MAX, filterOutArchivedSites };
export type { FilterResult };
