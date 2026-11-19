import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * delete-mail-draft moves one draft to Deleted Items, the one command that
 * puts a message there. It refuses a message that is not a draft. It moves the
 * draft (it does not DELETE it), so the answer has the new id, and the undo
 * moves the draft back to Drafts.
 */

const SEED: ReadonlyArray<StoredMessage> = [
  { id: 'd1', parentFolderId: 'drafts-id', changeKey: 'd1-ck', isDraft: true },
  { id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', isDraft: false },
  { id: 'd2', parentFolderId: 'deleted-id', changeKey: 'd2-ck', isDraft: true },
];

const writes = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

describe('delete-mail-draft', () => {
  it('is in the destructive class, with remove-mail-draft-attachment, the To Do commands that delete or move, delete-calendar-draft, and the Planner commands that delete', () => {
    expect(
      Object.entries(commands)
        .filter(([, command]) => command.meta.effect === 'delete')
        .map(([name]) => name)
    ).toEqual([
      'delete-mail-draft',
      'remove-mail-draft-attachment',
      'move-todo-task',
      'delete-todo-task',
      'apply-todo-ops',
      'delete-todo-task-step',
      'remove-todo-task-attachment',
      'delete-calendar-draft',
      'delete-planner-bucket',
      'delete-planner-task',
    ]);
  });

  it('shows the draft and its folder on a dry-run, and moves nothing', async () => {
    const mailbox = createMailStoreFake(SEED);
    const dry = await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd1' }, { dryRun: true });
    expect(dry['changes']).toEqual([
      { action: 'delete-draft', target: '/me/messages/d1/move', before: { messageId: 'd1', folderId: 'drafts-id', folder: 'Drafts', isDraft: 'true' } },
    ]);
    expect(writes(mailbox.calls)).toEqual([]);
  });

  it('moves the draft to Deleted Items with its new id, and the undo moves it back to Drafts', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd1' }, { journal });
    expect(run).toMatchObject({
      done: 1,
      runId: RUN_ID,
      changes: [{ after: { messageId: expect.stringMatching(/^d1-m\d+$/), folderId: 'deleted-id', folder: 'Deleted Items' } }],
    });
    expect(writes(mailbox.calls)).toEqual(['POST /me/messages/d1/move']);
    expect(statusesOf(await undoWrite(mailbox.graph, journal, false))).toEqual(['will-undo']);
    const report = await undoWrite(mailbox.graph, journal, true);
    expect(report.ok ? report.value.changes[0]?.now : report).toEqual({ messageId: expect.stringMatching(/^d1-m\d+$/), folderId: 'drafts-id' });
    expect([...mailbox.messages.values()].filter((m) => m.parentFolderId === 'drafts-id')).toHaveLength(1);
  });

  it('refuses a message that is not a draft, also on a dry-run', async () => {
    const mailbox = createMailStoreFake(SEED);
    for (const dryRun of [true, false]) {
      expect((await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'm1' }, { dryRun }))['error']).toEqual({
        code: 'not_a_draft',
        source: 'validation',
        message: 'The message m1 is not a draft, so delete-mail-draft does not delete it. This CLI deletes only unsent drafts. Check the id with get-mail-message.',
      });
    }
    expect(writes(mailbox.calls)).toEqual([]);
  });

  it('passes over a draft that is in Deleted Items already, and a draft that is gone', async () => {
    const mailbox = createMailStoreFake(SEED);
    const again = await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd2' });
    expect((again['skipped'] as ReadonlyArray<{ reason: string }>)[0]?.reason).toBe('The draft is in Deleted Items already.');
    const gone = await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd9' });
    expect(gone['done']).toBe(0);
    expect(writes(mailbox.calls)).toEqual([]);
  });

  it('passes over in the undo a draft that left Deleted Items, or that is no longer a draft', async () => {
    for (const fields of [{ parentFolderId: 'inbox-id' }, { isDraft: false }]) {
      const mailbox = createMailStoreFake(SEED);
      const journal = createRunJournalFake();
      const run = await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd1' }, { journal });
      mailbox.edit((run['changes'] as ReadonlyArray<{ after: { messageId: string } }>)[0]?.after.messageId ?? '', fields);
      expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('changed:')]);
    }
  });

  it('stops when Deleted Items cannot be read', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('GET /me/mailFolders/deleteditems', { type: 'api_error', status: 503, message: 'Busy.' });
    expect((await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'd1' }))['error']).toMatchObject({ message: 'Busy.' });
  });
});
