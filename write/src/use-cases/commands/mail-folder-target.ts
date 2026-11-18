import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphId, parseGraphId } from '../../domain/graph-id.ts';
import type { FolderReader, MailFolder } from './mail-folder-reader.ts';
import { isNotThere } from './mail-folder-reader.ts';

/*
 * A mail folder as a person names it (a well-known name, a path of display
 * names such as `Inbox/Projects`, a display name at the top, or an id), and the
 * guard on the folder that a move puts mail in. A move into Deleted Items,
 * Recoverable Items, Junk Email or Outbox, or into a folder in one of them, is
 * a delete, a junk report or a send, so the guard walks up the parents of the
 * target and refuses it. delete-mail-draft is the one command that puts a
 * message (a draft) in Deleted Items, and it does not use this guard.
 */

/** The folders that a move never puts mail in, by well-known name. Some mailboxes do not show each Recoverable Items folder. */
export const REFUSED_FOLDERS: ReadonlyArray<string> = [
  'deleteditems',
  'recoverableitemsroot',
  'recoverableitemsdeletions',
  'recoverableitemspurges',
  'recoverableitemsversions',
  'junkemail',
  'outbox',
];

const WELL_KNOWN: ReadonlySet<string> = new Set([
  ...REFUSED_FOLDERS,
  'inbox',
  'archive',
  'drafts',
  'sentitems',
  'msgfolderroot',
  'conversationhistory',
  'scheduled',
  'searchfolders',
  'syncissues',
  'clutter',
]);

// The well-known names are ids too: each passes the id check.
const wellKnown = (name: string): GraphId => graphId(name);

// A path such as `Inbox/Projects/Alpha`: its first name is a well-known name or a folder at the top.
const resolvePath = async (folders: FolderReader, names: ReadonlyArray<string>): Promise<Result<MailFolder, GraphError>> => {
  const [first = '', ...rest] = names;
  const lower = first.toLowerCase();
  let current = WELL_KNOWN.has(lower) ? await folders.folder(wellKnown(lower)) : await topFolder(folders, first);
  for (const name of rest) current = current.ok ? await folders.child(current.value, name) : current;
  return current;
};

const topFolder = async (folders: FolderReader, name: string): Promise<Result<MailFolder, GraphError>> => {
  const root = await folders.folder(wellKnown('msgfolderroot'));
  return root.ok ? folders.child(root.value, name) : root;
};

// An id, else a display name at the top.
const resolveIdOrName = async (folders: FolderReader, spec: string): Promise<Result<MailFolder, GraphError>> => {
  const id = parseGraphId(spec, 'folder');
  if (!id.ok) return topFolder(folders, spec);
  const byId = await folders.folder(id.value);
  return !byId.ok && isNotThere(byId.error) ? topFolder(folders, spec) : byId;
};

const noFolder = (what: string): GraphError => ({ type: 'validation_error', code: 'folder_not_found', message: `${what}: name a mail folder.` });

/** The folder that `spec` names. `what` names the param in a refusal, for example `--to`. */
export const resolveMailFolder = async (folders: FolderReader, spec: string, what: string): Promise<Result<MailFolder, GraphError>> => {
  const trimmed = spec.trim();
  if (trimmed === '') return err(noFolder(what));
  if (WELL_KNOWN.has(trimmed.toLowerCase())) return folders.folder(wellKnown(trimmed.toLowerCase()));
  if (!trimmed.includes('/')) return resolveIdOrName(folders, trimmed);
  return resolvePath(
    folders,
    trimmed.split('/').filter((name) => name !== '')
  );
};

// The ids of the refused folders that this mailbox has.
const refusedIds = async (folders: FolderReader): Promise<Result<ReadonlyMap<string, MailFolder>, GraphError>> => {
  const found = new Map<string, MailFolder>();
  for (const name of REFUSED_FOLDERS) {
    const folder = await folders.folder(wellKnown(name));
    if (folder.ok) found.set(folder.value.id, folder.value);
    else if (!isNotThere(folder.error)) return folder;
  }
  return found.size === 0 ? err(guardBlind) : ok(found);
};

const guardBlind: GraphError = {
  type: 'validation_error',
  code: 'folder_guard_failed',
  message: 'The folders that a move never puts mail in (Deleted Items, Junk Email, Outbox) cannot be read, so the move is refused.',
};

const placeOf = (target: MailFolder, refusedFolder: MailFolder): string => (target.id === refusedFolder.id ? 'is' : `is in "${refusedFolder.displayName}", which is`);

const refused = (what: string, target: MailFolder, refusedFolder: MailFolder): GraphError => ({
  type: 'validation_error',
  code: 'folder_refused',
  message: `${what}: "${target.displayName}" ${placeOf(target, refusedFolder)} a folder that no move puts mail in. A move into Deleted Items, Recoverable Items, Junk Email or Outbox, or into a folder in one of them, is a delete, a junk report or a send. To delete a draft, use delete-mail-draft.`,
});

// The deepest folder tree that the walk follows; Outlook allows far fewer levels.
const MAX_DEPTH = 32;

/** The folder `target`, or a refusal when it or a parent of it is a refused folder. */
export const refuseFolderInRefused = async (folders: FolderReader, target: MailFolder, what: string): Promise<Result<MailFolder, GraphError>> => {
  const ids = await refusedIds(folders);
  if (!ids.ok) return ids;
  let current: MailFolder = target;
  for (let depth = 0; depth < MAX_DEPTH; depth += 1) {
    const hit = ids.value.get(current.id);
    if (hit !== undefined) return err(refused(what, target, hit));
    const parentId = parseGraphId(current.parentFolderId, 'mail folder');
    // The root has no parent that the user can read: the walk ends there.
    const parent = parentId.ok && parentId.value !== current.id ? await folders.folder(parentId.value) : undefined;
    if (parent === undefined || (!parent.ok && isNotThere(parent.error))) return ok(target);
    if (!parent.ok) return parent;
    current = parent.value;
  }
  return ok(target);
};

/** The folder that `spec` names, refused when it is or is in a folder that no move puts mail in. */
export const moveTarget = async (folders: FolderReader, spec: string, what: string): Promise<Result<MailFolder, GraphError>> => {
  const target = await resolveMailFolder(folders, spec, what);
  return target.ok ? refuseFolderInRefused(folders, target.value, what) : target;
};
