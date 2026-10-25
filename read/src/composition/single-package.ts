/*
 * The read package `@ask-marcel/office-read` as the core shells see it: the
 * read commands, the lifecycle entries, its bin and its npm name. The write
 * commands are in @ask-marcel/office-write (`write-package.ts`). The shells
 * (CLI builder, MCP gateway, run-registry-command, docs builders) import no
 * registry; this module and `build-deps.ts` are where the real one is wired in.
 * Read's `bun run docs:gen` (it writes read/COMMANDS.md), read's build, and the
 * root `bun run docs:gen` and `bun run check:docs` read it through `--package`
 * (the name `commandPackage`).
 */
import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import { PACKAGE_NAMES } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { commands } from '../use-cases/commands/index.ts';
import { LIFECYCLE_ENTRIES } from '../use-cases/commands/lifecycle-entries.ts';
import { SINGLE_PACKAGE_NAMES } from './single-package-names.ts';

export const commandPackage: CommandPackage<Command<ReadGraph>> = {
  registry: commands,
  lifecycle: LIFECYCLE_ENTRIES,
  // The CLI adds the read counts that the registry gives.
  description:
    'Microsoft Graph CLI, read commands only: the write commands (mail drafts and PDF conversions) are in @ask-marcel/office-write. The CLI cannot create or modify calendar items (there is no send-mail / send-draft / create-event / upload-file command). Safe default for LLM autonomy.',
  exampleCommand: 'list-mail-messages',
  ...SINGLE_PACKAGE_NAMES,
  packageName: PACKAGE_NAMES.read,
};
