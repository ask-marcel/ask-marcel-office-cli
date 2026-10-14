/*
 * The single package's CLI: the core CLI builder (`cli-shell.ts`) given the
 * single package (`single-package.ts`) and the auth lifecycle commands
 * (`auth-cli.ts`). Its deps are the ones `main.ts` and the real-registry
 * suites have always passed.
 */
import type { Command } from 'commander';
import type { AuthCliDeps } from './auth-cli.ts';
import { addAuthCommands } from './auth-cli.ts';
import type { CliShellDeps } from './cli-shell.ts';
import { buildCliShell } from './cli-shell.ts';
import { commandPackage } from './single-package.ts';

type BuildCliDeps = Omit<CliShellDeps, 'commandPackage' | 'addLifecycleCommands'> & Omit<AuthCliDeps, 'names'>;

const buildCli = (deps: BuildCliDeps): Command =>
  buildCliShell({ ...deps, commandPackage, addLifecycleCommands: (program, output) => addAuthCommands(program, output, { ...deps, names: commandPackage }) });

export { buildCli };
export type { BuildCliDeps };
