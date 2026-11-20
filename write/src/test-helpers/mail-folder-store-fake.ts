import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { FOLDERS } from './mail-store-fake.ts';

/**
 * A mailbox whose folder tree can change, for the tests of create-mail-folder.
 * It starts with the folders of `FOLDERS` and answers:
 *
 * - `GET /me/mailFolders/{id or well-known name}`: the folder with its counts, or 404.
 * - `GET /me/mailFolders/{id}/childFolders?...`: the children of the folder.
 * - `POST /me/mailFolders/{id}/childFolders`: makes the folder `folder-<n>`, or 409 when the parent has a folder of that name (no case).
 * - `DELETE /me/mailFolders/{id}`: removes the folder.
 *
 * `calls` logs each request as `<METHOD> <path>`. `edit` changes a folder as
 * Outlook would (a new name, a message in it). `failures` makes a request fail:
 * the key is `<METHOD> <path>` (the path without its query).
 */
export type TreeFolder = {
  readonly id: string;
  readonly displayName: string;
  readonly parentFolderId: string;
  readonly wellKnownName?: string;
  readonly totalItemCount: number;
};

export type MailFolderStoreFake = {
  readonly graph: WriteGraph;
  readonly folders: Map<string, TreeFolder>;
  readonly calls: string[];
  readonly failures: Map<string, GraphError>;
  readonly edit: (id: string, fields: Partial<TreeFolder>) => void;
};

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

const exists: GraphError = { type: 'api_error', status: 409, code: 'ErrorFolderExists', message: 'A folder with the specified name already exists.' };

const pathOf = (path: string): string => path.split('?')[0] ?? '';

export const createMailFolderStoreFake = (): MailFolderStoreFake => {
  const folders = new Map<string, TreeFolder>(FOLDERS.map((folder) => [folder.id, { ...folder, totalItemCount: 0 }]));
  const calls: string[] = [];
  const failures = new Map<string, GraphError>();
  let made = 0;
  const find = (key: string): TreeFolder | undefined => [...folders.values()].find((folder) => folder.id === key || folder.wellKnownName === key);
  const children = (id: string): ReadonlyArray<TreeFolder> => [...folders.values()].filter((folder) => folder.parentFolderId === id);
  const shown = (folder: TreeFolder): Record<string, unknown> => ({
    id: folder.id,
    displayName: folder.displayName,
    parentFolderId: folder.parentFolderId,
    childFolderCount: children(folder.id).length,
    totalItemCount: folder.totalItemCount,
  });
  const logged = (method: string, path: string): GraphError | undefined => {
    calls.push(`${method} ${path}`);
    return failures.get(`${method} ${pathOf(path)}`);
  };
  const get = async (path: string): Promise<Result<unknown, GraphError>> => {
    const failed = logged('GET', path);
    if (failed !== undefined) return err(failed);
    const [, , , key = '', child] = pathOf(path).split('/');
    const folder = find(key);
    if (folder === undefined) return err(notFound);
    return ok(child === 'childFolders' ? { value: children(folder.id).map(shown) } : shown(folder));
  };
  const post = async (path: string, body: unknown): Promise<Result<unknown, GraphError>> => {
    const failed = logged('POST', path);
    if (failed !== undefined) return err(failed);
    const parent = find(pathOf(path).split('/')[3] ?? '');
    const name = String((body as { displayName?: unknown }).displayName);
    if (parent === undefined) return err(notFound);
    if (children(parent.id).some((folder) => folder.displayName.toLowerCase() === name.toLowerCase())) return err(exists);
    made += 1;
    const folder: TreeFolder = { id: `folder-${made}`, displayName: name, parentFolderId: parent.id, totalItemCount: 0 };
    folders.set(folder.id, folder);
    return ok({ ...shown(folder), unreadItemCount: 0, isHidden: false });
  };
  const remove = async (path: string): Promise<Result<unknown, GraphError>> => {
    const failed = logged('DELETE', path);
    if (failed !== undefined) return err(failed);
    const id = pathOf(path).split('/')[3] ?? '';
    return folders.delete(id) ? ok(undefined) : err(notFound);
  };
  return {
    graph: fakeWriteGraph({ get, post, delete: remove }),
    folders,
    calls,
    failures,
    edit: (id, fields) => {
      const folder = folders.get(id);
      if (folder !== undefined) folders.set(id, { ...folder, ...fields });
    },
  };
};
