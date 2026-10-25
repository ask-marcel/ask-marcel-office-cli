/*
 * The write package `@ask-marcel/office-write` as the core shells see it: the
 * write registry, its lifecycle entries, its own description, its bin and its
 * npm name. The write CLI (`write-program.ts`), the write MCP server
 * (`write-server.ts`) and the docs builders read it. Write's `bun run docs:gen`
 * (it writes write/COMMANDS.md), write's build and write's `bun run
 * check:docs` read it through `--package` (the name `commandPackage`).
 */
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import { PACKAGE_NAMES } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { commands } from '../use-cases/commands/write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from '../use-cases/commands/write-lifecycle-entries.ts';
import { WRITE_PACKAGE_NAMES } from './write-package-names.ts';

// The shell gives this text to commander as it is (it fills no bin), so it names
// packages, never a bin. The CLI adds the writes and what each leaves behind.
const WRITE_DESCRIPTION =
  'Microsoft Graph write commands for agents: mail drafts and the conversion of attachments to PDF. It sends no mail and cannot create or modify calendar items: a draft stays unsent until you send it. The read commands are in @ask-marcel/office-read, and the sign-in is in @ask-marcel/office-auth.';

export const commandPackage: CommandPackage<Command<WriteGraph>> = {
  registry: commands,
  lifecycle: WRITE_LIFECYCLE_ENTRIES,
  description: WRITE_DESCRIPTION,
  exampleCommand: 'create-mail-draft',
  ...WRITE_PACKAGE_NAMES,
  packageName: PACKAGE_NAMES.write,
};
