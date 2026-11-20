import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailFolderStoreFake } from '../../test-helpers/mail-folder-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * create-mail-folder on a small mailbox: the dry-run shows the parent and the
 * name and makes nothing, the run makes the folder with one call and shows its
 * id, and the undo deletes the folder only while it is empty, with its name,
 * in its parent.
 */

const writes = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

describe('the dry-run of create-mail-folder', () => {
  it('shows the top of the mailbox as the parent when no --parent is given, and makes nothing', async () => {
    const mailbox = createMailFolderStoreFake();
    expect(await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients' }, { dryRun: true })).toEqual({
      dryRun: true,
      command: 'create-mail-folder',
      changes: [
        {
          action: 'create-mail-folder',
          target: '/me/mailFolders/root-id/childFolders',
          before: { parentFolderId: 'root-id', parent: 'Top of Information Store', name: 'Clients' },
        },
      ],
    });
    expect(writes(mailbox.calls)).toEqual([]);
  });

  it('shows the folder of a --parent path, and the name with no space around it', async () => {
    const mailbox = createMailFolderStoreFake();
    const answer = await runWrite(mailbox.graph, 'create-mail-folder', { name: '  Beta  ', parent: 'Inbox/Projects' }, { dryRun: true });
    expect(answer['changes']).toEqual([
      { action: 'create-mail-folder', target: '/me/mailFolders/projects-id/childFolders', before: { parentFolderId: 'projects-id', parent: 'Projects', name: 'Beta' } },
    ]);
  });

  it('takes a well-known name, a folder name at the top and a folder id as the parent', async () => {
    const mailbox = createMailFolderStoreFake();
    const parents = [];
    for (const parent of ['ARCHIVE', 'reports', 'q1-id']) {
      const answer = await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients', parent }, { dryRun: true });
      parents.push((answer['changes'] as ReadonlyArray<{ before: { parent: string } }>)[0]?.before.parent);
    }
    expect(parents).toEqual(['Archive', 'Reports', 'Q1']);
  });
});

describe('the run of create-mail-folder', () => {
  it('makes the folder with one call, and shows the id of the new folder', async () => {
    const mailbox = createMailFolderStoreFake();
    const journal = createRunJournalFake();
    const answer = await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients', parent: 'inbox' }, { journal });
    expect(writes(mailbox.calls)).toEqual(['POST /me/mailFolders/inbox-id/childFolders']);
    expect(answer).toEqual({
      done: 1,
      skipped: [],
      failed: [],
      runId: RUN_ID,
      changes: [
        {
          action: 'create-mail-folder',
          target: '/me/mailFolders/inbox-id/childFolders',
          before: { parentFolderId: 'inbox-id', parent: 'Inbox', name: 'Clients' },
          after: { folderId: 'folder-1', name: 'Clients', parentFolderId: 'inbox-id' },
        },
      ],
    });
    expect(mailbox.folders.get('folder-1')).toMatchObject({ displayName: 'Clients', parentFolderId: 'inbox-id' });
    expect(journal.records.size).toBe(1);
  });

  it('makes a folder at the top of the mailbox when no --parent is given', async () => {
    const mailbox = createMailFolderStoreFake();
    await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients' });
    expect(mailbox.folders.get('folder-1')?.parentFolderId).toBe('root-id');
  });
});

describe('the undo of create-mail-folder', () => {
  const made = async (): Promise<{ mailbox: ReturnType<typeof createMailFolderStoreFake>; journal: ReturnType<typeof createRunJournalFake> }> => {
    const mailbox = createMailFolderStoreFake();
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients', parent: 'inbox' }, { journal });
    return { mailbox, journal };
  };

  it('shows the delete first, then deletes the empty folder', async () => {
    const { mailbox, journal } = await made();
    expect(statusesOf(await undoWrite(mailbox.graph, journal, false))).toEqual(['will-undo']);
    expect(mailbox.folders.has('folder-1')).toBe(true);
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.folders.has('folder-1')).toBe(false);
    expect(writes(mailbox.calls)).toEqual(['POST /me/mailFolders/inbox-id/childFolders', 'DELETE /me/mailFolders/folder-1']);
  });

  it('passes over a folder that has a message, a child folder, a new name or a new parent', async () => {
    const changes: ReadonlyArray<(mailbox: ReturnType<typeof createMailFolderStoreFake>) => void> = [
      (mailbox) => mailbox.edit('folder-1', { totalItemCount: 1 }),
      (mailbox) => mailbox.folders.set('child', { id: 'child', displayName: 'Child', parentFolderId: 'folder-1', totalItemCount: 0 }),
      (mailbox) => mailbox.edit('folder-1', { displayName: 'Customers' }),
      (mailbox) => mailbox.edit('folder-1', { parentFolderId: 'archive-id' }),
    ];
    for (const change of changes) {
      const { mailbox, journal } = await made();
      change(mailbox);
      expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([
        'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
      ]);
      expect(mailbox.folders.has('folder-1')).toBe(true);
    }
  });

  it('says that a folder deleted after the run is missing', async () => {
    const { mailbox, journal } = await made();
    mailbox.folders.delete('folder-1');
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['missing: The item is no longer there, so there is nothing to undo.']);
  });

  it('refuses a folder id of the journal that is not a Graph id before it goes into a path', async () => {
    const mailbox = createMailFolderStoreFake();
    const journal = createRunJournalFake();
    journal.records.set(RUN_ID, {
      version: 1,
      runId: RUN_ID,
      command: 'create-mail-folder',
      startedAt: '2026-11-18T09:05:07.000Z',
      changes: [
        {
          action: 'create-mail-folder',
          target: '/x',
          before: {},
          after: {},
          undo: { kind: 'delete-created-mail-folder', data: { folderId: '../send', name: 'Clients', parentFolderId: 'inbox-id' } },
        },
      ],
    });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([
      'failed: The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.',
    ]);
    expect(mailbox.calls).toEqual([]);
  });
});
