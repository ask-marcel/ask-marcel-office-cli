import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';

/**
 * A mailbox with folders behind a fake write graph, for the tests of the mail
 * state commands. It answers the requests that these commands make:
 *
 * - `GET /me/mailFolders/{id or well-known name}`: the folder, or 404.
 * - `GET /me/mailFolders/{id}/childFolders?...`: the children of the folder.
 * - `GET /me/mailFolders/{id}/messages?$filter=conversationId eq '...'`: the messages of a thread in the folder.
 * - `GET /me/messages/{id}?...`: the message, or 404.
 * - `POST /me/messages/{id}/move`: moves the message; it gets a new id (`<id>-m<n>`) and a new changeKey.
 * - `PATCH /me/messages/{id}`: merges the fields and gives a new changeKey.
 *
 * `calls` logs each request as `<METHOD> <path>`. `edit` changes a message as
 * Outlook would. `failures` makes a request fail: the key is `<METHOD> <path>`
 * (the path without its query), the value the error.
 */
export type StoredFolder = { readonly id: string; readonly displayName: string; readonly parentFolderId: string; readonly wellKnownName?: string };

export type StoredMessage = Record<string, unknown> & { readonly id: string; readonly parentFolderId: string; readonly changeKey: string };

export type MailStoreFake = {
  readonly graph: WriteGraph;
  readonly messages: Map<string, StoredMessage>;
  readonly calls: string[];
  readonly failures: Map<string, GraphError>;
  readonly edit: (id: string, fields: Record<string, unknown>) => void;
};

/** The folders of a small mailbox: the root, the well-known folders, `Inbox/Projects/Alpha`, `Deleted Items/Old` and `Reports/Q1`. */
export const FOLDERS: ReadonlyArray<StoredFolder> = [
  { id: 'root-id', displayName: 'Top of Information Store', parentFolderId: 'above-root', wellKnownName: 'msgfolderroot' },
  { id: 'inbox-id', displayName: 'Inbox', parentFolderId: 'root-id', wellKnownName: 'inbox' },
  { id: 'archive-id', displayName: 'Archive', parentFolderId: 'root-id', wellKnownName: 'archive' },
  { id: 'drafts-id', displayName: 'Drafts', parentFolderId: 'root-id', wellKnownName: 'drafts' },
  { id: 'sent-id', displayName: 'Sent Items', parentFolderId: 'root-id', wellKnownName: 'sentitems' },
  { id: 'deleted-id', displayName: 'Deleted Items', parentFolderId: 'root-id', wellKnownName: 'deleteditems' },
  { id: 'junk-id', displayName: 'Junk Email', parentFolderId: 'root-id', wellKnownName: 'junkemail' },
  { id: 'outbox-id', displayName: 'Outbox', parentFolderId: 'root-id', wellKnownName: 'outbox' },
  { id: 'recoverable-id', displayName: 'Recoverable Items', parentFolderId: 'above-root', wellKnownName: 'recoverableitemsroot' },
  { id: 'projects-id', displayName: 'Projects', parentFolderId: 'inbox-id' },
  { id: 'alpha-id', displayName: 'Alpha', parentFolderId: 'projects-id' },
  { id: 'old-id', displayName: 'Old', parentFolderId: 'deleted-id' },
  { id: 'reports-id', displayName: 'Reports', parentFolderId: 'root-id' },
  { id: 'q1-id', displayName: 'Q1', parentFolderId: 'reports-id' },
];

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

const segments = (path: string): ReadonlyArray<string> => (path.split('?')[0] ?? '').split('/');

const findFolder = (key: string): StoredFolder | undefined => FOLDERS.find((folder) => folder.id === key || folder.wellKnownName === key);

const conversationOf = (path: string): string => /conversationId%20eq%20'([^']*)'/.exec(path)?.[1] ?? '';

export const createMailStoreFake = (seed: ReadonlyArray<StoredMessage> = []): MailStoreFake => {
  const messages = new Map(seed.map((message) => [message.id, message]));
  const calls: string[] = [];
  const failures = new Map<string, GraphError>();
  let version = 1;
  const store = (message: StoredMessage): StoredMessage => {
    version += 1;
    const next = { ...message, changeKey: `ck-${version}` };
    messages.set(next.id, next);
    return next;
  };
  const failure = (method: string, path: string): GraphError | undefined => failures.get(`${method} ${path.split('?')[0] ?? ''}`);
  const folderAnswer = (parts: ReadonlyArray<string>, path: string): Result<unknown, GraphError> => {
    const folder = findFolder(parts[3] ?? '');
    if (folder === undefined) return err(notFound);
    if (parts[4] === 'childFolders') return ok({ value: FOLDERS.filter((child) => child.parentFolderId === folder.id) });
    if (parts[4] === 'messages') return ok({ value: [...messages.values()].filter((m) => m.parentFolderId === folder.id && m['conversationId'] === conversationOf(path)) });
    return ok({ id: folder.id, displayName: folder.displayName, parentFolderId: folder.parentFolderId });
  };
  const get = async (path: string): Promise<Result<unknown, GraphError>> => {
    calls.push(`GET ${path}`);
    const failed = failure('GET', path);
    if (failed !== undefined) return err(failed);
    const parts = segments(path);
    if (parts[2] === 'mailFolders') return folderAnswer(parts, path);
    const message = messages.get(parts[3] ?? '');
    return message === undefined ? err(notFound) : ok(structuredClone(message));
  };
  const post = async (path: string, body: unknown): Promise<Result<unknown, GraphError>> => {
    calls.push(`POST ${path}`);
    const failed = failure('POST', path);
    if (failed !== undefined) return err(failed);
    const message = messages.get(segments(path)[3] ?? '');
    const target = findFolder(String((body as { destinationId?: unknown }).destinationId));
    if (message === undefined || target === undefined) return err(notFound);
    messages.delete(message.id);
    return ok(structuredClone(store({ ...message, id: `${message.id.split('-m')[0] ?? ''}-m${version}`, parentFolderId: target.id })));
  };
  const patch = async (path: string, body: unknown): Promise<Result<unknown, GraphError>> => {
    calls.push(`PATCH ${path}`);
    const failed = failure('PATCH', path);
    if (failed !== undefined) return err(failed);
    const message = messages.get(segments(path)[3] ?? '');
    return message === undefined ? err(notFound) : ok(structuredClone(store({ ...message, ...(body as Record<string, unknown>) })));
  };
  return {
    graph: fakeWriteGraph({ get, post, patch }),
    messages,
    calls,
    failures,
    edit: (id, fields) => {
      const message = messages.get(id);
      if (message !== undefined) store({ ...message, ...fields });
    },
  };
};
