/*
 * Public library API of the read package (dist/index.js and dist/index.d.ts),
 * F3 of the package split:
 *
 * - `commands`: the read commands, each on the read graph alone;
 * - `createReadGraph(tokens)`: the read graph (every read tier, the chat
 *   substrate and a POST that reaches only the two query endpoints) over a
 *   token source;
 * - `createEnvThenHelperTokenSource()`: the tier variables of this process,
 *   else the token helper of @ask-marcel/office-auth;
 * - `createReadClient()`: the two above together.
 *
 * This entry imports no logger, so dist/index.js carries no winston. Core is
 * private and never published: the build inlines its code, and its types reach
 * the published declarations only through re-exports such as the ones below
 * (docs/plans/2026-10-01-package-split.md, Library API per package).
 */

import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import type { Command as CoreCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

export type { Result } from '@ask-marcel/office-core/domain/result.ts';
export { createReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
export type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
export type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';

/** A read command: its meta and its run on the read graph. */
export type Command = CoreCommand<ReadGraph>;

export { commands } from './composition/read-commands.ts';
export { createEnvThenHelperTokenSource } from './composition/read-token-source.ts';
export { createReadClient } from './composition/read-client.ts';
