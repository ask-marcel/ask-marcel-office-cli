import type { GraphClient } from '../../infra/graph-client.ts';

/**
 * The Microsoft Search index's security-trimmed match count for one entity type
 * (e.g. `driveItem` ≈ the number of files + folders the signed-in user can access).
 * A single size-1 `POST /search/query` reads `hitsContainers[0].total`. Best-effort:
 * returns `undefined` if the query fails or the index reports no numeric total, so
 * callers can surface the estimate when available without it ever failing the command.
 */

type SearchTotalResponse = { readonly value?: ReadonlyArray<{ readonly hitsContainers?: ReadonlyArray<{ readonly total?: unknown }> }> };

const searchIndexTotal = async (graph: GraphClient, entityType: string, queryString = '*'): Promise<number | undefined> => {
  const r = await graph.post('/search/query', { requests: [{ entityTypes: [entityType], query: { queryString }, size: 1 }] });
  if (!r.ok) return undefined;
  const total = (r.value as SearchTotalResponse).value?.[0]?.hitsContainers?.[0]?.total;
  return typeof total === 'number' ? total : undefined;
};

export { searchIndexTotal };
