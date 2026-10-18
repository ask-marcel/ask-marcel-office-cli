/*
 * The single package `ask-marcel-office-cli` as the core shells see it: the
 * whole registry, the lifecycle entries, its bin and its npm name. The shells
 * (CLI builder, MCP gateway, run-registry-command, docs builders) import no
 * registry; this module and `build-deps.ts` are where the real one is wired in.
 * `bun run docs:gen` and `bun run check:docs` read it through `--package`.
 * After the package split each package has its own (plan step 13).
 */
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import { commands } from '../use-cases/commands/index.ts';
import { LIFECYCLE_ENTRIES } from '../use-cases/commands/lifecycle-entries.ts';
import { SINGLE_PACKAGE_NAMES } from './single-package-names.ts';

export const commandPackage: CommandPackage = {
  registry: commands,
  lifecycle: LIFECYCLE_ENTRIES,
  // The CLI adds the writes and the read counts that the registry gives.
  description:
    'Microsoft Graph CLI. The CLI cannot create or modify calendar items (there is no send-mail / send-draft / create-event / upload-file command). Safe default for LLM autonomy.',
  exampleCommand: 'list-mail-messages',
  ...SINGLE_PACKAGE_NAMES,
  packageName: 'ask-marcel-office-cli',
};
