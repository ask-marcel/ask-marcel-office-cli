/*
 * Public library API of @ask-marcel/office-write (dist/index.js and
 * dist/index.d.ts), E4 of the package split:
 *
 * - `commands`: the 7 write commands, each on the write graph alone;
 * - `createWriteGraph(tokens)`: the write graph (the basic tier: get, post,
 *   patch, put, delete) over a token source;
 * - `createEnvThenHelperTokenSource()`: the tier variables of this process,
 *   else the token helper of @ask-marcel/office-auth;
 * - `createWriteClient()`: the two above together, the same write graph that
 *   the write bin builds.
 *
 * This entry imports no logger, so dist/index.js carries no winston. Core is
 * private and never published: the build inlines its code, and its types reach
 * the published declarations only through re-exports such as the ones below
 * (docs/plans/2026-10-01-package-split.md, Library API per package).
 */

import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import type { Command as CoreCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

export type { Result } from '@ask-marcel/office-core/domain/result.ts';
export { createWriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
export type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
export type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';

/** A write command: its meta and its run on the write graph. */
export type Command = CoreCommand<WriteGraph>;

export { commands } from './use-cases/commands/write-commands.ts';
export { createEnvThenHelperTokenSource } from './composition/write-token-source.ts';
export { createWriteClient } from './composition/write-client.ts';
