import { describe, expect, it } from 'bun:test';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * move-mail-message on a small mailbox: the dry-run shows the current folder of
 * each message, the run moves each message with one call and shows its id
 * before and after, and the undo moves each one back. A message that is not
 * there, that is in the target already, or that moved between the plan and the
 * write, is passed over and reported.
 */

const message = (id: string, folder: string, fields: Record<string, unknown> = {}): StoredMessage => ({
  id,
  parentFolderId: folder,
  changeKey: `${id}-ck`,
  isRead: false,
  ...fields,
});

const SEED = [message('m1', 'inbox-id'), message('m2', 'inbox-id'), message('m3', 'archive-id'), message('d1', 'drafts-id', { isDraft: true })];

const moves = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

describe('the dry-run of move-mail-message', () => {
  it('shows each message with its current folder, passes over a message in the target and one that is gone, and moves nothing', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1, d1,m3,gone', to: 'archive' }, { dryRun: true });
    expect(answer).toEqual({
      dryRun: true,
      command: 'move-mail-message',
      changes: [
        { action: 'move-message', target: '/me/messages/m1/move', before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox' } },
        { action: 'move-message', target: '/me/messages/d1/move', before: { messageId: 'd1', folderId: 'drafts-id', folder: 'Drafts' } },
      ],
      skipped: [
        { target: '/me/messages/m3/move', before: { messageId: 'm3', folderId: 'archive-id', folder: 'Archive' }, reason: 'The message is in "Archive" already.' },
        {
          target: '/me/messages/gone/move',
          before: { messageId: 'gone' },
          reason: 'The message is no longer there: it moved to another folder (a move gives it a new id), or it was deleted.',
        },
      ],
    });
    expect(moves(mailbox.calls)).toEqual([]);
  });

  it('reads each folder once for a batch', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1,m2', to: 'archive' }, { dryRun: true });
    expect(mailbox.calls.filter((call) => call.startsWith('GET /me/mailFolders/inbox-id?'))).toHaveLength(1);
  });
});

describe('the run of move-mail-message', () => {
  it('moves each message with one call, and shows the id before and after', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1,m2', to: 'Archive' }, { journal });
    expect(moves(mailbox.calls)).toEqual(['POST /me/messages/m1/move', 'POST /me/messages/m2/move']);
    expect(answer).toMatchObject({ done: 2, skipped: [], failed: [], runId: RUN_ID });
    const [first] = answer['changes'] as ReadonlyArray<Record<string, unknown>>;
    expect(first).toEqual({
      action: 'move-message',
      target: '/me/messages/m1/move',
      before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox' },
      after: { messageId: expect.stringMatching(/^m1-m\d+$/), folderId: 'archive-id', folder: 'Archive' },
    });
    expect([...mailbox.messages.values()].filter((m) => m.parentFolderId === 'archive-id')).toHaveLength(3);
  });

  it('takes the ids of --ids-file, one on each line, with the ids of --message-id, each once', async () => {
    const mailbox = createMailStoreFake(SEED);
    const files = createLocalFilesFake({ 'ids.txt': 'm1\r\n\n  m2  \nm1\n' });
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm2', idsFile: 'ids.txt', to: 'archive' }, { files });
    expect(answer['done']).toBe(2);
    expect(moves(mailbox.calls)).toEqual(['POST /me/messages/m2/move', 'POST /me/messages/m1/move']);
  });

  it('passes over a message that moved or changed between the plan and the write, and moves the others', async () => {
    const mailbox = createMailStoreFake(SEED);
    const planned = await commands['move-mail-message']?.plan?.(mailbox.graph, { messageId: 'm1,m2', to: 'archive' });
    mailbox.edit('m2', { isRead: true });
    const applied = planned?.ok === true ? await planned.value.apply() : undefined;
    expect(applied?.ok === true ? applied.value.value : applied).toEqual({
      done: 1,
      failed: [],
      skipped: [
        {
          target: '/me/messages/m2/move',
          before: { messageId: 'm2', folderId: 'inbox-id', folder: 'Inbox' },
          reason: 'The message moved or changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
        },
      ],
    });
    expect(moves(mailbox.calls)).toEqual(['POST /me/messages/m1/move']);
  });

  it('passes over a message that is gone between the plan and the write', async () => {
    const mailbox = createMailStoreFake(SEED);
    const planned = await commands['move-mail-message']?.plan?.(mailbox.graph, { messageId: 'm1', to: 'archive' });
    mailbox.messages.delete('m1');
    const applied = planned?.ok === true ? await planned.value.apply() : undefined;
    expect(applied?.ok === true ? (applied.value.value as { skipped: ReadonlyArray<{ reason: string }> }).skipped.map((item) => item.reason) : applied).toEqual([
      'The message is no longer there: it moved to another folder (a move gives it a new id), or it was deleted.',
    ]);
  });

  it('keeps the moves made in the journal when a move in the middle fails, so the undo reverses them', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('POST /me/messages/m2/move', { type: 'api_error', status: 503, message: 'Service unavailable.' });
    const journal = createRunJournalFake();
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1,m2,d1', to: 'archive' }, { journal });
    expect(answer).toMatchObject({ done: 2, runId: RUN_ID, failed: [{ target: '/me/messages/m2/move', reason: 'Service unavailable.' }] });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone', 'undone']);
    expect([...mailbox.messages.values()].map((m) => m.parentFolderId).toSorted((a, b) => a.localeCompare(b))).toEqual(['archive-id', 'drafts-id', 'inbox-id', 'inbox-id']);
  });

  it('answers the error when no move was made', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('POST /me/messages/m1/move', { type: 'api_error', status: 503, message: 'Service unavailable.' });
    const journal = createRunJournalFake();
    expect(await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }, { journal })).toEqual({
      error: expect.objectContaining({ message: 'Service unavailable.' }),
    });
    expect(journal.records.size).toBe(0);
  });

  it('records no run when each message is in the target already', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm3', to: 'archive' }, { journal });
    expect([answer['done'], answer['runId'], journal.records.size]).toEqual([0, undefined, 0]);
  });
});

describe('the undo of move-mail-message', () => {
  it('shows the moves back, then moves each message back to its folder, with its new id', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }, { journal });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, false))).toEqual(['will-undo']);
    const report = await undoWrite(mailbox.graph, journal, true);
    expect(report.ok ? report.value.changes[0]?.now : report).toEqual({ messageId: expect.stringMatching(/^m1-m\d+$/), folderId: 'inbox-id' });
    expect([...mailbox.messages.values()].filter((m) => m.parentFolderId === 'inbox-id')).toHaveLength(2);
  });

  it('passes over a message that is no longer in the folder the run put it in, and one that is gone', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1,m2', to: 'archive' }, { journal });
    const [first, second] = (answer['changes'] as ReadonlyArray<{ after: { messageId: string } }>).map((change) => change.after.messageId);
    mailbox.edit(first ?? '', { parentFolderId: 'sent-id' });
    mailbox.messages.delete(second ?? '');
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
      'missing: The item is no longer there, so there is nothing to undo.',
    ]);
    expect(moves(mailbox.calls)).toEqual(['POST /me/messages/m1/move', 'POST /me/messages/m2/move']);
  });

  it('never moves a message back into a refused folder', async () => {
    const mailbox = createMailStoreFake([message('j1', 'junk-id')]);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'j1', to: 'inbox' }, { journal });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([expect.stringContaining('failed: undo: "Junk Email" is a folder that no move puts mail in.')]);
  });

  it('refuses an id of the journal that is not a Graph id before it goes into a path', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    journal.records.set(RUN_ID, {
      version: 1,
      runId: RUN_ID,
      command: 'move-mail-message',
      startedAt: '2026-11-18T09:05:07.000Z',
      changes: [{ action: 'move-message', target: '/x', before: {}, after: {}, undo: { kind: 'move-message-back', data: { messageId: '../send', folderId: 'a', backTo: 'b' } } }],
    });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['failed: The run journal names an id that is not a Graph id, so this undo does not run.']);
    expect(mailbox.calls).toEqual([]);
  });
});
