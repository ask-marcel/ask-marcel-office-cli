import type { GraphClient } from '../../infra/graph-client.ts';
import { searchIndexTotal } from './search-index-total.ts';

/**
 * Opt-in per-entry file count: scope the Microsoft Search `driveItem` index to
 * ONE site/drive by its webUrl (KQL `path:`) and read the security-trimmed
 * `total` (files + folders). One Search query per entry, so this is behind a
 * flag and the fan-out is chunked + capped. Double-quotes in the URL are
 * stripped before interpolation so they cannot break out of the KQL phrase.
 */

const FILE_COUNT_CHUNK = 10;
const FILE_COUNT_MAX = 200;

const countForUrl = async (graph: GraphClient, webUrl: string): Promise<number | undefined> =>
  webUrl === '' ? undefined : searchIndexTotal(graph, 'driveItem', `path:"${webUrl.replaceAll('"', '')}"`);

const countsForUrls = async (graph: GraphClient, urls: ReadonlyArray<string>, max: number, chunk: number): Promise<ReadonlyArray<number | undefined>> => {
  const out: Array<number | undefined> = [];
  for (let start = 0; start < urls.length; start += chunk) {
    const slice = urls.slice(start, start + chunk);
    const done = await Promise.all(slice.map((url, j) => (start + j < max ? countForUrl(graph, url) : Promise.resolve<number | undefined>(undefined))));
    out.push(...done);
  }
  return out;
};

type FileCountOptions = { readonly max?: number; readonly chunk?: number };

// Return a copy of `entries` with `estimatedFileCount` merged onto each object
// entry whose webUrl yields a count; entries past the cap or without a webUrl
// are returned unchanged.
const addEstimatedFileCounts = async (
  graph: GraphClient,
  entries: ReadonlyArray<unknown>,
  webUrlOf: (entry: unknown) => string | undefined,
  options?: FileCountOptions
): Promise<ReadonlyArray<unknown>> => {
  const max = options?.max ?? FILE_COUNT_MAX;
  const chunk = options?.chunk ?? FILE_COUNT_CHUNK;
  const counts = await countsForUrls(
    graph,
    entries.map((e) => webUrlOf(e) ?? ''),
    max,
    chunk
  );
  return entries.map((entry, i) => {
    const count = counts[i];
    return count !== undefined && entry !== null && typeof entry === 'object' ? { ...entry, estimatedFileCount: count } : entry;
  });
};

export { addEstimatedFileCounts };
