/*
 * The auth package `@ask-marcel/office-auth` as the core shells see it: no
 * registry (it has no Graph command), its six lifecycle entries, its own
 * description, its bin and its npm name. The auth CLI (`auth-program.ts`) and
 * the docs builders read it.
 */
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import { PACKAGE_NAMES } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { AUTH_LIFECYCLE_ENTRIES } from '../use-cases/commands/auth-lifecycle-entries.ts';
import { AUTH_PACKAGE_NAMES } from './auth-package-names.ts';

const AUTH_DESCRIPTION =
  'Microsoft 365 sign-in for the @ask-marcel/office-read and @ask-marcel/office-write packages. `login` signs in through the Teams web client and keeps the tokens in ~/.ask-marcel/token-cache.json. `status` shows the tokens, and `logout` removes them. The read and write packages get their tokens from this package.';

export const authPackage: CommandPackage = {
  registry: {},
  lifecycle: AUTH_LIFECYCLE_ENTRIES,
  description: AUTH_DESCRIPTION,
  exampleCommand: 'login',
  ...AUTH_PACKAGE_NAMES,
  packageName: PACKAGE_NAMES.auth,
};
