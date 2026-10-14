/*
 * The single package's CLI, the module `main.ts` and the real-registry suites
 * import. The builder is in `cli-shell.ts`, which package-split step 13 turns
 * into the core CLI builder.
 */
export { buildCli } from './cli-shell.ts';
export type { BuildCliDeps } from './cli-shell.ts';
