import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';

/*
 * The reads of the mail folders that the mail state commands need: a folder by
 * its id or its well-known name, and a child folder by its display name. One
 * reader is made for each plan or undo, and it keeps each folder that it read,
 * so a batch reads a folder once.
 */

export type MailFolder = { readonly id: GraphId; readonly displayName: string; readonly parentFolderId: string };

export type FolderReader = {
  /** The folder with this id or well-known name (`inbox`, `deleteditems`). */
  readonly folder: (key: GraphId) => Promise<Result<MailFolder, GraphError>>;
  /** The child of the folder `parent` whose display name is `name`, compared with no case. */
  readonly child: (parent: MailFolder, name: string) => Promise<Result<MailFolder, GraphError>>;
};

const SELECT = '$select=id,displayName,parentFolderId';

const folderSchema = z.object({ id: z.string(), displayName: z.string(), parentFolderId: z.string().optional() });

const unreadable: GraphError = { type: 'validation_error', message: 'Graph gave a mail folder in a shape that this CLI cannot read.' };

// A folder of a Graph answer, its id checked before it goes into a path.
const folderOf = (value: unknown): Result<MailFolder, GraphError> => {
  const parsed = folderSchema.safeParse(value);
  if (!parsed.success) return err(unreadable);
  const id = parseGraphId(parsed.data.id, 'mail folder');
  return id.ok ? ok({ id: id.value, displayName: parsed.data.displayName, parentFolderId: parsed.data.parentFolderId ?? '' }) : id;
};

const childrenSchema = z.object({ value: z.array(z.unknown()) });

const noChild = (parent: MailFolder, name: string): GraphError => ({
  type: 'validation_error',
  code: 'folder_not_found',
  message: `No folder named "${name}" in "${parent.displayName}". Run list-mail-folders or list-mail-child-folders to see the folder names, or give the folder id.`,
});

const findChild = (value: unknown, parent: MailFolder, name: string): Result<MailFolder, GraphError> => {
  const list = childrenSchema.safeParse(value);
  if (!list.success) return err(unreadable);
  const match = list.data.value.find((entry) => folderSchema.safeParse(entry).data?.displayName.toLowerCase() === name.toLowerCase());
  return match === undefined ? err(noChild(parent, name)) : folderOf(match);
};

export const createFolderReader = (reads: GraphReads): FolderReader => {
  const known = new Map<string, MailFolder>();
  const folder = async (key: GraphId): Promise<Result<MailFolder, GraphError>> => {
    const hit = known.get(key);
    if (hit !== undefined) return ok(hit);
    const read = await reads.get(`/me/mailFolders/${key}?${SELECT}`);
    const found = read.ok ? folderOf(read.value) : read;
    if (found.ok) known.set(key, found.value).set(found.value.id, found.value);
    return found;
  };
  const child = async (parent: MailFolder, name: string): Promise<Result<MailFolder, GraphError>> => {
    // 500 a page: the default page of a folder listing is ten (LESSONS 2026-09-26), and hidden folders count too.
    const read = await reads.get(`/me/mailFolders/${parent.id}/childFolders?$top=500&includeHiddenFolders=true&${SELECT}`);
    return read.ok ? findChild(read.value, parent, name) : read;
  };
  return { folder, child };
};

/** Whether a Graph error says that the folder or the message is not there (404), or that Graph does not know the id (400). */
export const isNotThere = (error: GraphError): boolean => error.type === 'api_error' && (error.status === 404 || error.status === 400);
