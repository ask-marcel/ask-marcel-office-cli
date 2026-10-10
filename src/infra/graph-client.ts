import type { AuthManager } from '../infra/auth.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { createAuthManagerTokenSource } from './auth-token-source.ts';
import type { FetchFn, GraphError } from './graph-request.ts';
import type { ReadGraph } from './read-graph.ts';
import { createReadGraph } from './read-graph.ts';
import type { WriteGraph } from './write-graph.ts';
import { createWriteGraph } from './write-graph.ts';

/*
 * The single package's full client: the read graph and the write graph on one
 * object, for the composition root and for a library caller. The package
 * split drops this type: read builds only the read graph, write only the write
 * graph (docs/plans/2026-10-01-package-split.md, D9).
 */
type GraphClient = ReadGraph & WriteGraph;

// Every bearer comes from the token source, which knows the tiers; the two
// graphs only sign requests with them. The write graph's POST, which reaches
// any path, is the full client's.
const createTokenSourceGraphClient = (tokens: TokenSource, fetchFn: FetchFn = globalThis.fetch): GraphClient => ({
  ...createReadGraph(tokens, fetchFn),
  ...createWriteGraph(tokens, fetchFn),
});

// The default: the in-process auth ladder, or a library caller's own manager.
const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => createTokenSourceGraphClient(createAuthManagerTokenSource(auth), fetchFn);

export { createGraphClient, createTokenSourceGraphClient };
export type { FetchFn, GraphClient, GraphError };
