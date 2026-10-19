/*
 * The library helper of the write package: the write graph, signed by the
 * tier variables, else by the token helper of the auth package. Its own module
 * with no logger, so the library bundle carries no winston.
 */
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { createWriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { createEnvThenHelperTokenSource } from './write-token-source.ts';

export const createWriteClient = (): WriteGraph => createWriteGraph(createEnvThenHelperTokenSource());
