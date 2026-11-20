import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailFolderStoreFake } from '../../test-helpers/mail-folder-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * Each refusal of create-mail-folder, as the caller gets it: its message, its
 * code and its source (`validation`: a wrong input, which a retry does not
 * cure). Each refusal comes before any write. A fault of Graph is the error
 * of the run, and the journal records no run.
 */

const refusal = async (params: Record<string, string>, mailbox = createMailFolderStoreFake()): Promise<unknown> => {
  const answer = await runWrite(mailbox.graph, 'create-mail-folder', params);
  expect(mailbox.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
  return answer['error'];
};

const BAD_CHARACTER = '--name: a folder name cannot have "/", "\\" or a control character. --parent reads "/" as the step from a folder to its child folder.';

describe('the refusals of a wrong name', () => {
  it('refuse a name that is empty or only spaces', async () => {
    for (const name of ['', '   ']) {
      expect(await refusal({ name })).toEqual({ code: 'invalid_folder_name', source: 'validation', message: '--name: give the name of the new folder.' });
    }
  });

  it('take a name of 255 characters and refuse a name of 256', async () => {
    const mailbox = createMailFolderStoreFake();
    expect(await runWrite(mailbox.graph, 'create-mail-folder', { name: 'a'.repeat(255) }, { dryRun: true })).toMatchObject({ dryRun: true });
    expect(await refusal({ name: 'a'.repeat(256) })).toEqual({
      code: 'invalid_folder_name',
      source: 'validation',
      message: '--name: a folder name has at most 255 characters. This name has 256.',
    });
  });

  it('refuse a slash, a backslash and a control character in the name', async () => {
    for (const name of ['Clients/2026', 'Clients\\2026', 'Clients\t2026', 'Clients\u007f']) {
      expect([name, await refusal({ name })]).toEqual([name, { code: 'invalid_folder_name', source: 'validation', message: BAD_CHARACTER }]);
    }
  });

  it('refuse a name that the parent has already, with no case', async () => {
    expect(await refusal({ name: 'PROJECTS', parent: 'inbox' })).toEqual({
      code: 'folder_exists',
      source: 'validation',
      message: '--name: "Inbox" has a folder named "Projects" already. Give another name, or use that folder.',
    });
    expect(await refusal({ name: 'inbox' })).toMatchObject({ code: 'folder_exists' });
  });

  it('say which required param is missing', async () => {
    expect(await refusal({ parent: 'inbox' })).toEqual({ source: 'validation', message: expect.stringContaining('--name') });
  });
});

describe('the refusals of a wrong parent', () => {
  it('refuse Deleted Items, Junk Email, Outbox, Recoverable Items and a folder in one of them', async () => {
    expect(await refusal({ name: 'Clients', parent: 'Deleted Items/Old' })).toEqual({
      code: 'folder_refused',
      source: 'validation',
      message:
        '--parent: "Old" is in "Deleted Items", which is a folder where this CLI makes no folder. Deleted Items, Recoverable Items, Junk Email and Outbox, and each folder in one of them, are for deleted mail, junk mail or mail to send.',
    });
    for (const parent of ['deleteditems', 'junkemail', 'outbox', 'recoverableitemsroot']) {
      expect([parent, await refusal({ name: 'Clients', parent })]).toEqual([parent, expect.objectContaining({ code: 'folder_refused', source: 'validation' })]);
    }
    expect(await refusal({ name: 'Clients', parent: 'junkemail' })).toMatchObject({ message: expect.stringContaining('"Junk Email" is a folder where this CLI makes no folder.') });
  });

  it('refuse the create when the guard cannot read the refused folders', async () => {
    const mailbox = createMailFolderStoreFake();
    for (const id of ['deleted-id', 'recoverable-id', 'junk-id', 'outbox-id']) mailbox.folders.delete(id);
    expect(await refusal({ name: 'Clients', parent: 'inbox' }, mailbox)).toEqual({
      code: 'folder_guard_failed',
      source: 'validation',
      message: 'The folders where this CLI makes no folder (Deleted Items, Junk Email, Outbox) cannot be read, so the create is refused.',
    });
  });

  it('refuse an empty parent and a parent that is not there', async () => {
    expect(await refusal({ name: 'Clients', parent: ' ' })).toEqual({ code: 'folder_not_found', source: 'validation', message: '--parent: name a mail folder.' });
    expect(await refusal({ name: 'Clients', parent: 'Inbox/Nowhere' })).toMatchObject({ code: 'folder_not_found', source: 'validation' });
  });
});

describe('the faults of Graph', () => {
  it('give the error of a read of the children that is not a 404, before any write', async () => {
    const mailbox = createMailFolderStoreFake();
    mailbox.failures.set('GET /me/mailFolders/inbox-id/childFolders', { type: 'api_error', status: 503, message: 'Service unavailable.' });
    expect(await refusal({ name: 'Clients', parent: 'inbox' }, mailbox)).toMatchObject({ message: 'Service unavailable.' });
  });

  it('give the error of the create and record no run', async () => {
    const mailbox = createMailFolderStoreFake();
    mailbox.failures.set('POST /me/mailFolders/inbox-id/childFolders', { type: 'api_error', status: 409, message: 'A folder with the specified name already exists.' });
    const journal = createRunJournalFake();
    expect(await runWrite(mailbox.graph, 'create-mail-folder', { name: 'Clients', parent: 'inbox' }, { journal })).toEqual({
      error: expect.objectContaining({ message: 'A folder with the specified name already exists.' }),
    });
    expect(journal.records.size).toBe(0);
  });

  it('say that Graph made a folder whose id this CLI cannot read, so the run has no undo for it', async () => {
    const mailbox = createMailFolderStoreFake();
    const graph = { ...mailbox.graph, post: async () => ok({ displayName: 'Clients' }) };
    const journal = createRunJournalFake();
    expect(await runWrite(graph, 'create-mail-folder', { name: 'Clients' }, { journal })).toEqual({
      error: {
        source: 'validation',
        message: 'Graph made the folder but gave no id that this CLI can read, so the run has no undo for it. Run list-mail-folders to find it.',
      },
    });
    expect(journal.records.size).toBe(0);
  });
});
