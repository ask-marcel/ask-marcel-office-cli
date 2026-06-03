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

// `$expand=drive($select=quota)` rides on the SAME probe so each kept site gets
// its default library's `quota.used` (total bytes, recursive) for free — no extra
// call. The site $select still includes siteCollection, so archive detection is
// unaffected; a site with no default drive simply comes back without `drive`.
const PROBE_PATH = '?$select=id,webUrl,siteCollection&$expand=drive($select=quota)';
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
type ProbeOutcome = { readonly verdict: Verdict; readonly size?: number };

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

const driveSize = (body: unknown): number | undefined => {
  if (body === null || typeof body !== 'object') return undefined;
  const used = (body as { drive?: { quota?: { used?: unknown } } }).drive?.quota?.used;
  return typeof used === 'number' ? used : undefined;
};

const probe = async (graph: GraphClient, resource: unknown): Promise<ProbeOutcome> => {
  const url = siteWebUrl(resource);
  if (url !== undefined && isNonNavigableSiteUrl(url)) return { verdict: 'nonNavigable' };
  const id = siteId(resource);
  if (id === undefined) return { verdict: 'keep' };
  const r = await graph.get(`/sites/${id}${PROBE_PATH}`);
  if (r.ok) return { verdict: isArchivedSite(r.value) ? 'archived' : 'keep', size: driveSize(r.value) };
  if (errorIndicatesArchived(errorFields(r.error))) return { verdict: 'archived' };
  if (r.error.type === 'api_error' && r.error.status === 404) return { verdict: 'notFound' };
  return { verdict: 'error' };
};

const verdictsFor = async (graph: GraphClient, sites: ReadonlyArray<unknown>, probeMax: number, chunkSize: number): Promise<ReadonlyArray<ProbeOutcome>> => {
  const out: Array<ProbeOutcome> = [];
  for (let start = 0; start < sites.length; start += chunkSize) {
    const chunk = sites.slice(start, start + chunkSize);
    const done = await Promise.all(chunk.map((site, j) => (start + j < probeMax ? probe(graph, site) : Promise.resolve<ProbeOutcome>({ verdict: 'keep' }))));
    out.push(...done);
  }
  return out;
};

const countVerdict = (outcomes: ReadonlyArray<ProbeOutcome>, target: Verdict): number => outcomes.filter((o) => o.verdict === target).length;

// Merge the per-site size onto a kept site resource (only when known and the
// resource is an object); leaves the resource untouched otherwise.
const withSize = (resource: unknown, size: number | undefined): unknown =>
  size !== undefined && resource !== null && typeof resource === 'object' ? { ...resource, size } : resource;

const filterOutArchivedSites = async (graph: GraphClient, sites: ReadonlyArray<unknown>, options?: FilterOptions): Promise<FilterResult> => {
  const probeMax = options?.probeMax ?? ARCHIVE_PROBE_MAX;
  const chunkSize = options?.chunkSize ?? PROBE_CHUNK;
  const outcomes = await verdictsFor(graph, sites, probeMax, chunkSize);
  const value = sites
    .map((site, i) => ({ site, outcome: outcomes[i] }))
    .filter((e) => e.outcome?.verdict === 'keep' || e.outcome?.verdict === 'error')
    .map((e) => withSize(e.site, e.outcome?.size));
  return {
    value,
    archivedExcluded: countVerdict(outcomes, 'archived'),
    nonNavigableExcluded: countVerdict(outcomes, 'nonNavigable'),
    notFoundExcluded: countVerdict(outcomes, 'notFound'),
    probeErrors: countVerdict(outcomes, 'error'),
    probeTruncated: sites.length > probeMax,
  };
};

export { ARCHIVE_PROBE_MAX, filterOutArchivedSites };
export type { FilterResult };
