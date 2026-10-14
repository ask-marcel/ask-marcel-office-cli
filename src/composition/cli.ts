/*
 * The single package's CLI: the CLI builder (`cli-shell.ts`) given the auth
 * lifecycle commands (`auth-cli.ts`). Its deps are the ones `main.ts` and the real-registry
 * suites have always passed.
 */
import type { Command } from 'commander';
import type { AuthCliDeps } from './auth-cli.ts';
import { addAuthCommands } from './auth-cli.ts';
import type { CliShellDeps } from './cli-shell.ts';
import { buildCliShell } from './cli-shell.ts';

type BuildCliDeps = Omit<CliShellDeps, 'addLifecycleCommands'> & AuthCliDeps;

const buildCli = (deps: BuildCliDeps): Command => buildCliShell({ ...deps, addLifecycleCommands: (program, output) => addAuthCommands(program, output, deps) });

export { buildCli };
export type { BuildCliDeps };
