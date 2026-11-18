import { describe, expect, it } from 'bun:test';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * The guards of the moves. A move into Deleted Items, Recoverable Items, Junk
 * Email or Outbox, or into a folder in one of them, is refused, however the
 * folder is named: the guard walks up the parents of the target and compares
 * ids, not names. A folder is named by a well-known name, a path, a name at the
 * top or an id. An id from a param or a file goes into a Graph path only after
 * the id check.
 */

const SEED: ReadonlyArray<StoredMessage> = [{ id: 'm1', parentFolderId: 'inbox-id', changeKey: 'ck', isRead: false, conversationId: 'c1' }];

const errorOf = (answer: Record<string, unknown>): { readonly code?: string; readonly message?: string } => answer['error'] as { code?: string; message?: string };

const move = async (to: string, mailbox = createMailStoreFake(SEED)): Promise<Record<string, unknown>> => runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to });

describe('the target of a move', () => {
  for (const to of ['deleteditems', 'DeletedItems', 'Deleted Items', 'deleted-id', 'junkemail', 'Junk Email', 'outbox', 'recoverableitemsroot', 'recoverable-id']) {
    it(`refuses ${to}, and moves nothing`, async () => {
      const mailbox = createMailStoreFake(SEED);
      expect(errorOf(await move(to, mailbox)).code).toBe('folder_refused');
      expect(mailbox.calls.filter((call) => call.startsWith('POST'))).toEqual([]);
    });
  }

  it('refuses a folder in Deleted Items, named by its path or by its id, and names the folder it is in', async () => {
    expect(errorOf(await move('Deleted Items/Old')).message).toBe(
      '--to: "Old" is in "Deleted Items", which is a folder that no move puts mail in. A move into Deleted Items, Recoverable Items, Junk Email or Outbox, or into a folder in one of them, is a delete, a junk report or a send. To delete a draft, use delete-mail-draft.'
    );
    expect(errorOf(await move('old-id')).code).toBe('folder_refused');
  });

  it('says that the target is the refused folder itself', async () => {
    expect(errorOf(await move('Outbox')).message).toStartWith('--to: "Outbox" is a folder that no move puts mail in.');
  });

  it('takes a well-known name, a path, a name at the top and an id', async () => {
    for (const [to, folder] of [
      ['ARCHIVE', 'archive-id'],
      ['Inbox/Projects/Alpha', 'alpha-id'],
      ['/inbox//projects/', 'projects-id'],
      ['Sent Items', 'sent-id'],
      ['alpha-id', 'alpha-id'],
    ] as const) {
      const mailbox = createMailStoreFake(SEED);
      await move(to, mailbox);
      expect([to, [...mailbox.messages.values()][0]?.parentFolderId]).toEqual([to, folder]);
    }
  });

  it('refuses a name that no folder has, and an empty name', async () => {
    expect(errorOf(await move('Projects')).message).toBe(
      'No folder named "Projects" in "Top of Information Store". Run list-mail-folders or list-mail-child-folders to see the folder names, or give the folder id.'
    );
    expect(errorOf(await move('Inbox/Nope')).code).toBe('folder_not_found');
    expect(errorOf(await move('  ')).message).toBe('--to: name a mail folder.');
  });

  it('refuses the move when no refused folder can be read, so it never moves blind', async () => {
    const mailbox = createMailStoreFake(SEED);
    for (const name of ['deleteditems', 'recoverableitemsroot', 'recoverableitemsdeletions', 'recoverableitemspurges', 'recoverableitemsversions', 'junkemail', 'outbox'])
      mailbox.failures.set(`GET /me/mailFolders/${name}`, { type: 'api_error', status: 404, message: 'Not found.' });
    expect(errorOf(await move('archive', mailbox))).toMatchObject({
      code: 'folder_guard_failed',
      message: 'The folders that a move never puts mail in (Deleted Items, Junk Email, Outbox) cannot be read, so the move is refused.',
    });
  });

  it('stops on an error of Graph that is not a missing folder', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('GET /me/mailFolders/junkemail', { type: 'api_error', status: 503, message: 'Service unavailable.' });
    expect(errorOf(await move('archive', mailbox)).message).toBe('Service unavailable.');
    const walk = createMailStoreFake(SEED);
    walk.failures.set('GET /me/mailFolders/projects-id', { type: 'api_error', status: 503, message: 'Busy.' });
    expect(errorOf(await move('alpha-id', walk)).message).toBe('Busy.');
  });

  it('ends the walk at the root, whose parent cannot be read', async () => {
    const mailbox = createMailStoreFake(SEED);
    await move('archive', mailbox);
    expect(mailbox.calls).toContain('GET /me/mailFolders/above-root?$select=id,displayName,parentFolderId');
    expect(mailbox.calls.filter((call) => call.startsWith('POST'))).toEqual(['POST /me/messages/m1/move']);
  });
});

describe('the id check', () => {
  for (const id of ['m1/send', 'm1?x=1', "m1' or 1 eq 1", 'm1%2f..', 'm 1', 'x'.repeat(513)]) {
    it(`refuses the message id ${JSON.stringify(id.slice(0, 20))} before any call to Graph`, async () => {
      const mailbox = createMailStoreFake(SEED);
      const answer = await runWrite(mailbox.graph, 'mark-mail-message', { messageId: `m1,${id}`, read: 'true' });
      expect(errorOf(answer).code).toBe('invalid_id');
      expect(mailbox.calls).toEqual([]);
    });
  }

  it('takes an id of 512 characters', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'x'.repeat(512), read: 'true' });
    expect(mailbox.calls).toEqual([`GET /me/messages/${'x'.repeat(512)}?$select=parentFolderId,changeKey,isRead`]);
  });

  it('names the line of --ids-file that holds a bad id, and never quotes the content of the file', async () => {
    const files = createLocalFilesFake({ 'ids.txt': 'm1\n\nSECRET=abc/def\n' });
    const answer = await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { idsFile: 'ids.txt', read: 'true' }, { files });
    expect(errorOf(answer)).toMatchObject({
      code: 'invalid_id',
      message: '--ids-file line 3 is not a Graph id. Give one message id on each line, as list-mail-folder-messages returns it.',
    });
  });

  it('refuses a conversation id that could change the $filter', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answer = await runWrite(mailbox.graph, 'move-mail-thread', { conversationId: "c1') or (1 eq 1", from: 'inbox', to: 'archive' });
    expect(errorOf(answer).message).toStartWith('--conversation-id: "c1\') or (1 eq 1" is not a Graph id.');
    expect(mailbox.calls).toEqual([]);
  });

  it('refuses a run with no message, and a batch of more than 1000', async () => {
    expect(errorOf(await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { messageId: ' , ', read: 'true' })).message).toBe(
      'Give the messages: --message-id <id> (ids separated by commas) or --ids-file <file> (one id on each line).'
    );
    const many = Array.from({ length: 1001 }, (_, index) => `m${index}`).join(',');
    expect(errorOf(await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { messageId: many, read: 'true' })).message).toBe(
      '1001 messages: one run changes at most 1000. Divide the ids into more runs.'
    );
  });

  it('takes a batch of 1000 messages', async () => {
    const many = Array.from({ length: 1000 }, (_, index) => `m${index}`).join(',');
    const answer = await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { messageId: many, read: 'true' }, { dryRun: true });
    expect([(answer['changes'] as ReadonlyArray<unknown>).length, (answer['skipped'] as ReadonlyArray<unknown>).length]).toEqual([1, 999]);
  });

  it('shows no skipped list on a dry-run that passes over nothing', async () => {
    const answer = await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { messageId: 'm1', read: 'true' }, { dryRun: true });
    expect(Object.keys(answer)).toEqual(['dryRun', 'command', 'changes']);
  });

  it('refuses --ids-file when the caller gives no local file access, and a file that is not there', async () => {
    const graph = createMailStoreFake(SEED).graph;
    expect(errorOf(await runWrite(graph, 'move-mail-message', { idsFile: 'ids.txt', to: 'archive' })).code).toBe('local_files_unavailable');
    expect(errorOf(await runWrite(graph, 'move-mail-message', { idsFile: 'ids.txt', to: 'archive' }, { files: createLocalFilesFake() })).code).toBe('local_file_not_found');
  });
});
