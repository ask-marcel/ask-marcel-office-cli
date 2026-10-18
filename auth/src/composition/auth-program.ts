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

const AUTH_DESCRIPTION =
  'Microsoft 365 sign-in for the @ask-marcel/office-read and @ask-marcel/office-write packages. `login` signs in through the Teams web client and keeps the tokens in ~/.ask-marcel/token-cache.json. `status` shows the tokens, and `logout` removes them. The read and write packages get their tokens from this package.';

type BuildAuthCliDeps = Omit<LifecycleCliDeps, 'commandPackage' | 'description' | 'addLifecycleCommands'> & Omit<AuthCliDeps, 'names'>;

const buildAuthCli = (deps: BuildAuthCliDeps): Command =>
  buildLifecycleCli({
    ...deps,
    commandPackage: authPackage,
    description: AUTH_DESCRIPTION,
    addLifecycleCommands: (program, output) => addAuthCommands(program, output, { ...deps, names: authPackage }),
  });

export { buildAuthCli };
export type { BuildAuthCliDeps };
