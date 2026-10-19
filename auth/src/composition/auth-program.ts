/*
 * The CLI of the auth package: the core lifecycle CLI (`buildLifecycleCli`)
 * given the auth package (`auth-package.ts`) and the sign-in commands
 * (`auth-cli.ts`). It takes no graph: the auth package runs no Graph command.
 */
import type { Command } from 'commander';
import type { AuthCliDeps } from './auth-cli.ts';
import { addAuthCommands } from './auth-cli.ts';
import { authPackage } from './auth-package.ts';
import type { LifecycleCliDeps } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { buildLifecycleCli } from '@ask-marcel/office-core/composition/cli-shell.ts';

type BuildAuthCliDeps = Omit<LifecycleCliDeps, 'commandPackage' | 'addLifecycleCommands'> & Omit<AuthCliDeps, 'names'>;

const buildAuthCli = (deps: BuildAuthCliDeps): Command =>
  buildLifecycleCli({
    ...deps,
    commandPackage: authPackage,
    addLifecycleCommands: (program, output) => addAuthCommands(program, output, { ...deps, names: authPackage }),
  });

export { buildAuthCli };
export type { BuildAuthCliDeps };
