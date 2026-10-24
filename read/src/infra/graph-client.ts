import type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';
import type { FetchFn, GraphError } from '@ask-marcel/office-core/infra/graph-request.ts';
import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import { createReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';

/*
 * The root package's client: the read graph, for the composition root. The
 * root holds the read commands only, so its client has no write member (no
 * put, patch, delete or upload session, and a POST that reaches only the two
 * query endpoints). The write commands and the write graph are in
 * @ask-marcel/office-write (docs/plans/2026-10-01-package-split.md, D9).
 */
type GraphClient = ReadGraph;

// Every bearer comes from the token source, which knows the tiers; the read
// graph only signs requests with them.
const createTokenSourceGraphClient = (tokens: TokenSource, fetchFn: FetchFn = globalThis.fetch): GraphClient => createReadGraph(tokens, fetchFn);

export { createTokenSourceGraphClient };
export type { FetchFn, GraphClient, GraphError };
