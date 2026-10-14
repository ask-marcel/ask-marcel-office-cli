/*
 * The single package `ask-marcel-office-cli` as the core shells see it: the
 * whole registry, the lifecycle entries, its bin and its npm name.
 * run-registry-command and the docs builders take it as an argument instead of
 * importing the registry. After the package split each package has its own
 * (plan step 13).
 */
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { commands } from '../use-cases/commands/index.ts';
import { LIFECYCLE_ENTRIES } from '../use-cases/commands/lifecycle-entries.ts';

export const commandPackage: CommandPackage = {
  registry: commands,
  lifecycle: LIFECYCLE_ENTRIES,
  binName: 'ask-marcel-office',
  packageName: 'ask-marcel-office-cli',
};
