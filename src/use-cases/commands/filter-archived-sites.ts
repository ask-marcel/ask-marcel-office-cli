import { errorIndicatesArchived, isArchivedSite } from '../../domain/utilities/archive-status.ts';
import { isNonNavigableSiteUrl } from '../../domain/utilities/site-url-classifier.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Drop sites the signed-in user cannot actually open from a list of Microsoft
 * Graph `site` resources, so the search index's noise never reaches an LLM:
 *   - NON-NAVIGABLE by URL shape (add-in app domains, `/contentstorage/`
 *     SharePoint Embedded containers, `/_layouts/` system pages) — dropped
 *     deterministically from `webUrl`, no probe (`nonNavigableExcluded`).
 *   - ARCHIVED — each remaining site is probed (`GET /sites/{id}?$select=
 *     id,webUrl,siteCollection`); dropped when Graph reports it archived
 *     (`archiveStatus`) or fails with an archive/locked signal (423
 *     `resourceLocked`) (`archivedExcluded`).
 *   - NOT FOUND — a probe that 404s means the site does not resolve; dropped
 *     (`notFoundExcluded`).
 * A probe that fails for any OTHER reason KEEPS the site (we never lose an
 * accessible site to a transient error) and is counted in `probeErrors`.
 * Probing is capped and concurrency-bounded to avoid 429 throttling; sites past
 * the cap are kept unprobed and `probeTruncated` is set. Active personal
 * OneDrives are kept — they are not dropped by URL shape, only by a failing probe.
 */

const PROBE_PATH = '?$select=id,webUrl,siteCollection';
const ARCHIVE_PROBE_MAX = 250;
const PROBE_CHUNK = 15;

type FilterOptions = { readonly probeMax?: number; readonly chunkSize?: number };
type FilterResult = {
  readonly value: ReadonlyArray<unknown>;
  readonly archivedExcluded: number;
  readonly nonNavigableExcluded: number;
  readonly notFoundExcluded: number;
  readonly probeErrors: number;
  readonly probeTruncated: boolean;
};
type Verdict = 'keep' | 'archived' | 'nonNavigable' | 'notFound' | 'error';

const siteId = (resource: unknown): string | undefined => {
  if (resource === null || typeof resource !== 'object') return undefined;
  const id = (resource as { id?: unknown }).id;
  return typeof id === 'string' ? id : undefined;
};

const siteWebUrl = (resource: unknown): string | undefined => {
  if (resource === null || typeof resource !== 'object') return undefined;
  const url = (resource as { webUrl?: unknown }).webUrl;
  return typeof url === 'string' ? url : undefined;
};

const errorFields = (e: GraphError): { status?: number; code?: string; message?: string } => ({
  status: e.type === 'api_error' ? e.status : undefined,
  code: e.code,
  message: e.message,
});

const probe = async (graph: GraphClient, resource: unknown): Promise<Verdict> => {
  const url = siteWebUrl(resource);
  if (url !== undefined && isNonNavigableSiteUrl(url)) return 'nonNavigable';
  const id = siteId(resource);
  if (id === undefined) return 'keep';
  const r = await graph.get(`/sites/${id}${PROBE_PATH}`);
  if (r.ok) return isArchivedSite(r.value) ? 'archived' : 'keep';
  if (errorIndicatesArchived(errorFields(r.error))) return 'archived';
  if (r.error.type === 'api_error' && r.error.status === 404) return 'notFound';
  return 'error';
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

const countVerdict = (verdicts: ReadonlyArray<Verdict>, target: Verdict): number => verdicts.filter((v) => v === target).length;

const filterOutArchivedSites = async (graph: GraphClient, sites: ReadonlyArray<unknown>, options?: FilterOptions): Promise<FilterResult> => {
  const probeMax = options?.probeMax ?? ARCHIVE_PROBE_MAX;
  const chunkSize = options?.chunkSize ?? PROBE_CHUNK;
  const verdicts = await verdictsFor(graph, sites, probeMax, chunkSize);
  const value = sites.filter((_, i) => verdicts[i] === 'keep' || verdicts[i] === 'error');
  return {
    value,
    archivedExcluded: countVerdict(verdicts, 'archived'),
    nonNavigableExcluded: countVerdict(verdicts, 'nonNavigable'),
    notFoundExcluded: countVerdict(verdicts, 'notFound'),
    probeErrors: countVerdict(verdicts, 'error'),
    probeTruncated: sites.length > probeMax,
  };
};

export { ARCHIVE_PROBE_MAX, filterOutArchivedSites };
export type { FilterResult };
