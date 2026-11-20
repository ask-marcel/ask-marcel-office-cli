import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailFolderStoreFake } from '../../test-helpers/mail-folder-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { meta } from './create-mail-folder.ts';

/*
 * The docs of create-mail-folder, as the help, the docs pages and the MCP
 * tools show them, and the edges of its write and of its undo: an answer of
 * Graph that this CLI cannot read never goes into the journal or a path.
 */

describe('the docs of create-mail-folder', () => {
  it('pins the meta of create-mail-folder', () => {
    expect(meta).toStrictEqual({
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
          description:
            'The name of the new folder, at most 255 characters, with no "/", "\\" or control character. The parent cannot have a folder of this name already (no case).',
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
    });
  });
});

describe('the edges of create-mail-folder', () => {
  it('sends the name with no space around it as the display name', async () => {
    const mailbox = createMailFolderStoreFake();
    const bodies: unknown[] = [];
    const graph = {
      ...mailbox.graph,
      post: async (path: string, body: unknown) => {
        bodies.push(body);
        return mailbox.graph.post(path, body);
      },
    };
    await runWrite(graph, 'create-mail-folder', { name: ' Clients ' });
    expect(bodies).toEqual([{ displayName: 'Clients' }]);
  });

  it('refuses an id of the new folder that is not a Graph id, and records no run', async () => {
    const mailbox = createMailFolderStoreFake();
    const graph = { ...mailbox.graph, post: async () => ok({ id: 'folder/1', displayName: 'Clients' }) };
    const journal = createRunJournalFake();
    expect(await runWrite(graph, 'create-mail-folder', { name: 'Clients' }, { journal })).toMatchObject({
      error: { message: expect.stringContaining('Graph made the folder but gave no id that this CLI can read') },
    });
    expect(journal.records.size).toBe(0);
  });

  it('reads the name, the parent and the counts of the folder before its undo', async () => {
    const mailbox = createMailFolderStoreFake();
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients' }, { journal });
    await undoWrite(mailbox.graph, journal, false);
    expect(mailbox.calls.at(-1)).toBe('GET /me/mailFolders/folder-1?$select=displayName,parentFolderId,childFolderCount,totalItemCount');
  });

  it('passes over a folder that Graph gives in a shape that this CLI cannot read', async () => {
    const mailbox = createMailFolderStoreFake();
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients' }, { journal });
    const graph = { ...mailbox.graph, get: async () => ok({ displayName: 'Clients' }) };
    expect(statusesOf(await undoWrite(graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
    ]);
    expect(mailbox.folders.has('folder-1')).toBe(true);
  });
});
