/*
 * The library helper of the read package: the read graph, signed by the tier
 * variables, else by the token helper of the auth package. Its own module
 * with no logger, so the library bundle carries no winston.
 */
import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import { createReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import { createEnvThenHelperTokenSource } from './read-token-source.ts';

export const createReadClient = (): ReadGraph => createReadGraph(createEnvThenHelperTokenSource());
