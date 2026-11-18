import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The three commands that change a state field of a message: the read state
 * (mark-mail-message), the follow-up flag (set-mail-flag) and the categories
 * (set-mail-categories). Each shows the value before in a dry-run, writes it
 * with one PATCH, passes over a message that has the value already, and its
 * undo writes the value of before back only if the field still has the value
 * that the run set.
 */

const SEED: ReadonlyArray<StoredMessage> = [
  { id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', isRead: false, flag: { flagStatus: 'notFlagged' }, categories: ['Blue'] },
  { id: 'm2', parentFolderId: 'inbox-id', changeKey: 'm2-ck', isRead: true, flag: { flagStatus: 'notFlagged' }, categories: [] },
];

const patches = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => call.startsWith('PATCH'));

describe('mark-mail-message', () => {
  it('shows the read state before, marks only the messages that need it, and the undo marks them back', async () => {
    const mailbox = createMailStoreFake(SEED);
    const dry = await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1,m2', read: 'true' }, { dryRun: true });
    expect(dry['changes']).toEqual([{ action: 'mark-message', target: '/me/messages/m1', before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox', isRead: 'false' } }]);
    expect(dry['skipped']).toEqual([
      { target: '/me/messages/m2', before: { messageId: 'm2', folderId: 'inbox-id', folder: 'Inbox', isRead: 'true' }, reason: 'The message is read already.' },
    ]);
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1,m2', read: 'true' }, { journal });
    expect(run).toMatchObject({ done: 1, runId: RUN_ID, changes: [{ after: { messageId: 'm1', isRead: 'true' } }] });
    expect(mailbox.messages.get('m1')?.['isRead']).toBe(true);
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.messages.get('m1')?.['isRead']).toBe(false);
  });

  it('marks unread, and says so of a message that is unread already', async () => {
    const mailbox = createMailStoreFake(SEED);
    const run = await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1,m2', read: 'false' });
    expect([run['done'], (run['skipped'] as ReadonlyArray<{ reason: string }>)[0]?.reason]).toEqual([1, 'The message is unread already.']);
    expect(patches(mailbox.calls)).toEqual(['PATCH /me/messages/m2']);
  });

  it('passes over a message whose read state changed after the run, and opens no other message', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1', read: 'true' }, { journal });
    mailbox.edit('m1', { isRead: false });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('changed:')]);
    expect(patches(mailbox.calls)).toEqual(['PATCH /me/messages/m1']);
  });

  it('undoes a change of the read state although the changeKey changed after the run', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1', read: 'true' }, { journal });
    mailbox.edit('m1', { categories: ['Red'] });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
  });

  it('refuses a --read that is not true or false', async () => {
    expect((await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { messageId: 'm1', read: 'yes' }))['error']).toMatchObject({
      message: expect.stringContaining('read'),
    });
  });
});

describe('set-mail-flag', () => {
  it('flags a message with a due day, and the undo removes the flag', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm1', status: 'flagged', due: '2026-11-20' }, { journal });
    expect(run['changes']).toEqual([
      {
        action: 'set-flag',
        target: '/me/messages/m1',
        before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox', flag: 'notFlagged' },
        after: { messageId: 'm1', flag: 'flagged, due 2026-11-20' },
      },
    ]);
    const day = { dateTime: '2026-11-20T00:00:00', timeZone: 'UTC' };
    expect(mailbox.messages.get('m1')?.['flag']).toEqual({ flagStatus: 'flagged', startDateTime: day, dueDateTime: day });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.messages.get('m1')?.['flag']).toEqual({ flagStatus: 'notFlagged' });
  });

  it('marks a flag complete, and passes over a message that has that flag already', async () => {
    const mailbox = createMailStoreFake(SEED);
    expect((await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm1', status: 'complete' }))['done']).toBe(1);
    expect(mailbox.messages.get('m1')?.['flag']).toEqual({ flagStatus: 'complete' });
    const again = await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm1', status: 'complete' });
    expect((again['skipped'] as ReadonlyArray<{ reason: string }>)[0]?.reason).toBe('The message has the flag "complete" already.');
  });

  it('passes over a flag that changed after the run, and reads a message with no flag as not flagged', async () => {
    const mailbox = createMailStoreFake([{ id: 'm3', parentFolderId: 'inbox-id', changeKey: 'ck' }]);
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm3', status: 'flagged' }, { journal });
    expect((run['changes'] as ReadonlyArray<{ before: { flag: string } }>)[0]?.before.flag).toBe('notFlagged');
    mailbox.edit('m3', { flag: { flagStatus: 'flagged', dueDateTime: { dateTime: '2026-12-01T00:00:00.0000000', timeZone: 'UTC' } } });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('changed:')]);
  });

  it('refuses a due date with another status, and a day that does not exist', async () => {
    const graph = createMailStoreFake(SEED).graph;
    expect((await runWrite(graph, 'set-mail-flag', { messageId: 'm1', status: 'complete', due: '2026-11-20' }))['error']).toMatchObject({
      message: '--due: a due date goes with --status flagged only.',
    });
    for (const due of ['2026-02-30', '20261120', '2026-11-20T10:00']) {
      expect((await runWrite(graph, 'set-mail-flag', { messageId: 'm1', status: 'flagged', due }))['error']).toMatchObject({
        message: expect.stringContaining('must be a day in the form YYYY-MM-DD'),
      });
    }
    expect((await runWrite(graph, 'set-mail-flag', { messageId: 'm1/x', status: 'flagged' }))['error']).toMatchObject({ code: 'invalid_id' });
  });
});

describe('set-mail-categories', () => {
  it('replaces the categories, each once, and the undo gives the categories of before back', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'set-mail-categories', { messageId: 'm1', categories: ' Red , Project,Red,' }, { journal });
    expect(run['changes']).toEqual([
      {
        action: 'set-categories',
        target: '/me/messages/m1',
        before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox', categories: 'Blue' },
        after: { messageId: 'm1', categories: 'Project, Red' },
      },
    ]);
    expect(mailbox.messages.get('m1')?.['categories']).toEqual(['Red', 'Project']);
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.messages.get('m1')?.['categories']).toEqual(['Blue']);
  });

  it('removes all the categories with an empty --categories, and passes over a message with the same categories in another order', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'set-mail-categories', { messageId: 'm1', categories: '' });
    expect(mailbox.messages.get('m1')?.['categories']).toEqual([]);
    mailbox.edit('m2', { categories: ['B', 'A'] });
    const same = await runWrite(mailbox.graph, 'set-mail-categories', { messageId: 'm2', categories: 'A,B' });
    expect((same['skipped'] as ReadonlyArray<{ reason: string }>)[0]?.reason).toBe('The message has these categories already.');
  });

  it('passes over categories that changed after the run', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'set-mail-categories', { messageId: 'm1', categories: 'Red' }, { journal });
    mailbox.edit('m1', { categories: ['Red', 'Green'] });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('changed:')]);
    mailbox.messages.delete('m1');
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('missing:')]);
  });
});
