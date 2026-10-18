/*
 * The auth package `@ask-marcel/office-auth` as the core shells see it: no
 * registry (it has no Graph command), its six lifecycle entries, its bin and
 * its npm name. The auth CLI (`auth-program.ts`) and the docs builders read it.
 */
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import { PACKAGE_NAMES } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { AUTH_LIFECYCLE_ENTRIES } from '../use-cases/commands/auth-lifecycle-entries.ts';
import { AUTH_PACKAGE_NAMES } from './auth-package-names.ts';

export const authPackage: CommandPackage = {
  registry: {},
  lifecycle: AUTH_LIFECYCLE_ENTRIES,
  ...AUTH_PACKAGE_NAMES,
  packageName: PACKAGE_NAMES.auth,
};
