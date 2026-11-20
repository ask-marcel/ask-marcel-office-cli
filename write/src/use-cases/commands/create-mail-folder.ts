import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseMailFolderName } from '../../domain/mail-folder-name.ts';
import { createFolderOp } from './mail-folder-create.ts';
import type { FolderGuardWords } from './mail-folder-target.ts';
import { refuseFolderInRefused, resolveMailFolder } from './mail-folder-target.ts';
import type { FolderReader, MailFolder } from './mail-folder-reader.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Makes a mail folder, at the top of the mailbox or in the folder --parent.
 * Before any write, the plan refuses a parent that is, or is in, Deleted
 * Items, Recoverable Items, Junk Email or Outbox (the guard of the moves,
 * mail-folder-target.ts), and a name that the parent has already: Graph
 * refuses it too, and one read of the children finds it first.
 */

const schema = z.object({ name: z.string(), parent: z.string().optional() });

const CREATE_WORDS: FolderGuardWords = {
  refusedFolder: 'a folder where this CLI makes no folder',
  why: 'Deleted Items, Recoverable Items, Junk Email and Outbox, and each folder in one of them, are for deleted mail, junk mail or mail to send.',
  blind: 'The folders where this CLI makes no folder (Deleted Items, Junk Email, Outbox) cannot be read, so the create is refused.',
};

const taken = (parent: MailFolder, name: string): GraphError => ({
  type: 'validation_error',
  code: 'folder_exists',
  message: `--name: "${parent.displayName}" has a folder named "${name}" already. Give another name, or use that folder.`,
});

// The parent, when it has no folder of this name (no case).
const refuseTaken = async (folders: FolderReader, parent: MailFolder, name: string): Promise<Result<MailFolder, GraphError>> => {
  const child = await folders.child(parent, name);
  if (child.ok) return err(taken(parent, child.value.displayName));
  return child.error.code === 'folder_not_found' ? ok(parent) : child;
};

const parentOf = async (folders: FolderReader, spec: string | undefined, name: string): Promise<Result<MailFolder, GraphError>> => {
  const parent = await resolveMailFolder(folders, spec ?? 'msgfolderroot', '--parent');
  const allowed = parent.ok ? await refuseFolderInRefused(folders, parent.value, '--parent', CREATE_WORDS) : parent;
  return allowed.ok ? refuseTaken(folders, allowed.value, name) : allowed;
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const name = parseMailFolderName(parsed.value.name);
  if (!name.ok) return name;
  const parent = await parentOf(createFolderReader(reads), parsed.value.parent, name.value);
  return parent.ok ? ok(todoPlan({ ops: [createFolderOp(parent.value, name.value)] })) : parent;
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Create a mail folder, at the top of the mailbox or in the folder --parent. Refuses a parent that is, or is in, Deleted Items, Recoverable Items, Junk Email or Outbox, and a name that the parent has already. The undo deletes the folder while it is empty.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/mailFolders/<the id of {parent}, else of msgfolderroot>/childFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-post-childfolders',
  options: [
    {
      name: 'name',
      key: 'name',
      required: true,
      description: 'The name of the new folder, at most 255 characters, with no "/", "\\" or control character. The parent cannot have a folder of this name already (no case).',
    },
    {
      name: 'parent',
      key: 'parent',
      required: false,
      description:
        'The folder to make the new folder in: a well-known name (inbox, archive), a path of folder names from the top (`Inbox/Projects`), a folder name at the top, or a folder id from list-mail-folders. With no --parent, the folder goes at the top of the mailbox.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} create-mail-folder --name "Clients" --parent inbox --dry-run',
  bodyTemplate: "{ displayName: '{name}' }",
  effect: 'mail-state',
  undo: 'deletes the folder, only if it has no message and no child folder, and its name and its parent did not change',
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    '`{ done, skipped, failed, runId, changes: [{ action, target, before, after }] }`. `before` has the `parentFolderId`, the `parent` name and the `name` of the new folder. `after` has the `folderId` of the new folder, its `name` and its `parentFolderId`. With --dry-run: `{ dryRun: true, changes }`, and nothing changes.',
};

export { execute, meta, plan, schema };
