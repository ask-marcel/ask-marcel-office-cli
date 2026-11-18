import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { graphId } from '../../domain/graph-id.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * The edges of the mail state commands: the folders as Graph names them, the
 * answers that Graph gives in an odd shape, a write that fails, the undo data
 * of a journal that a person edited, and the `execute` of a library caller.
 */

const SEED: ReadonlyArray<StoredMessage> = [
  { id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', isRead: false },
  { id: 'm2', parentFolderId: 'inbox-id', changeKey: 'm2-ck', isRead: false },
];

const folderOfM1 = (mailbox: ReturnType<typeof createMailStoreFake>): string | undefined => [...mailbox.messages.values()].find((m) => m.id.startsWith('m1'))?.parentFolderId;

const errorOf = (answer: Record<string, unknown>): { readonly code?: string; readonly message?: string } => answer['error'] as { code?: string; message?: string };

// A graph that answers the paths that `answer` knows, and asks the mailbox for the others.
const answering = (graph: WriteGraph, answer: (path: string) => Result<unknown, GraphError> | undefined): WriteGraph => ({
  ...graph,
  get: async (path) => answer(path) ?? graph.get(path),
});

const WELL_KNOWN = [
  'deleteditems',
  'recoverableitemsroot',
  'recoverableitemsdeletions',
  'recoverableitemspurges',
  'recoverableitemsversions',
  'junkemail',
  'outbox',
  'inbox',
  'archive',
  'drafts',
  'sentitems',
  'msgfolderroot',
  'conversationhistory',
  'scheduled',
  'searchfolders',
  'syncissues',
  'clutter',
];

describe('the folders of a move', () => {
  it('reads each well-known name in lower case, whatever its case', async () => {
    for (const name of WELL_KNOWN) {
      const mailbox = createMailStoreFake(SEED);
      await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: name.toUpperCase() }, { dryRun: true });
      expect(mailbox.calls[0]).toBe(`GET /me/mailFolders/${name}?$select=id,displayName,parentFolderId`);
    }
  });

  it('reads each refused folder for the guard', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }, { dryRun: true });
    for (const name of WELL_KNOWN.slice(0, 7)) expect(mailbox.calls).toContain(`GET /me/mailFolders/${name}?$select=id,displayName,parentFolderId`);
  });

  it('finds a folder name with no case, in a listing of 500 that holds the hidden folders', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'SENT items' });
    expect(mailbox.calls).toContain('GET /me/mailFolders/root-id/childFolders?$top=500&includeHiddenFolders=true&$select=id,displayName,parentFolderId');
    expect(folderOfM1(mailbox)).toBe('sent-id');
    const path = createMailStoreFake(SEED);
    await runWrite(path.graph, 'move-mail-message', { messageId: 'm1', to: 'reports/q1' });
    expect(folderOfM1(path)).toBe('q1-id');
  });

  it('reads a name that has the form of an id as a name at the top when no folder has that id, and stops on another error', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'Reports' });
    expect(folderOfM1(mailbox)).toBe('reports-id');
    const unknownId = createMailStoreFake(SEED);
    unknownId.failures.set('GET /me/mailFolders/Reports', { type: 'api_error', status: 400, message: 'Id is malformed.' });
    await runWrite(unknownId.graph, 'move-mail-message', { messageId: 'm1', to: 'Reports' });
    expect(folderOfM1(unknownId)).toBe('reports-id');
    const busy = createMailStoreFake(SEED);
    busy.failures.set('GET /me/mailFolders/Reports', { type: 'api_error', status: 503, message: 'Busy.' });
    expect(errorOf(await runWrite(busy.graph, 'move-mail-message', { messageId: 'm1', to: 'Reports' })).message).toBe('Busy.');
  });

  it('reads a folder that it knows by its id once, also when it read it by its well-known name', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'Inbox/Projects' }, { dryRun: true });
    expect(mailbox.calls.filter((call) => call.startsWith('GET /me/mailFolders/inbox-id?'))).toEqual([]);
  });

  it('ends the walk of the parents at a folder that is its own parent, and at a loop of parents', async () => {
    const loop = (path: string): Result<unknown, GraphError> | undefined => {
      if (path.startsWith('/me/mailFolders/self-id?')) return ok({ id: 'self-id', displayName: 'Self', parentFolderId: 'self-id' });
      if (path.startsWith('/me/mailFolders/x-id?')) return ok({ id: 'x-id', displayName: 'X', parentFolderId: 'y-id' });
      return path.startsWith('/me/mailFolders/y-id?') ? ok({ id: 'y-id', displayName: 'Y', parentFolderId: 'x-id' }) : undefined;
    };
    for (const to of ['self-id', 'x-id']) {
      const mailbox = createMailStoreFake(SEED);
      const answer = await runWrite(answering(mailbox.graph, loop), 'move-mail-message', { messageId: 'm1', to }, { dryRun: true });
      expect(answer['changes']).toHaveLength(1);
      expect(mailbox.calls.filter((call) => call.startsWith('GET /me/mailFolders/self-id?'))).toEqual([]);
    }
  });

  it('refuses a folder that Graph gives in a shape it cannot read, and a listing it cannot read', async () => {
    const mailbox = createMailStoreFake(SEED);
    const odd = answering(mailbox.graph, (path) => (path.startsWith('/me/mailFolders/archive?') ? ok({ name: 'Archive' }) : undefined));
    expect(errorOf(await runWrite(odd, 'move-mail-message', { messageId: 'm1', to: 'archive' })).message).toBe('Graph gave a mail folder in a shape that this CLI cannot read.');
    const list = answering(mailbox.graph, (path) => (path.includes('/childFolders') ? ok({ folders: [] }) : undefined));
    expect(errorOf(await runWrite(list, 'move-mail-message', { messageId: 'm1', to: 'Reports' })).message).toBe('Graph gave a mail folder in a shape that this CLI cannot read.');
    const badId = answering(mailbox.graph, (path) => (path.startsWith('/me/mailFolders/archive?') ? ok({ id: 'a/b', displayName: 'Archive' }) : undefined));
    expect(errorOf(await runWrite(badId, 'move-mail-message', { messageId: 'm1', to: 'archive' })).code).toBe('invalid_id');
  });

  it('shows an empty folder name when the folder of a message cannot be read', async () => {
    const mailbox = createMailStoreFake([
      { id: 'm1', parentFolderId: 'lost-id', changeKey: 'ck' },
      { id: 'm2', parentFolderId: 'a/b', changeKey: 'ck' },
    ]);
    const answer = await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1,m2', to: 'archive' }, { dryRun: true });
    expect((answer['changes'] as ReadonlyArray<{ before: { folder: string } }>).map((change) => change.before.folder)).toEqual(['', '']);
  });
});

describe('the answers of Graph to the writes', () => {
  it('refuses a message that Graph gives in a shape it cannot read', async () => {
    const mailbox = createMailStoreFake(SEED);
    const odd = answering(mailbox.graph, (path) => (path.startsWith('/me/messages/m1?') ? ok({ id: 'm1' }) : undefined));
    expect(errorOf(await runWrite(odd, 'mark-mail-message', { messageId: 'm1', read: 'true' })).message).toBe('Graph gave the message m1 in a shape that this CLI cannot read.');
  });

  it('refuses a move whose answer has no new id, or a new id that is not an id', async () => {
    for (const answer of [{}, { id: 'a/b', parentFolderId: 'archive-id' }]) {
      const mailbox = createMailStoreFake(SEED);
      const graph = { ...mailbox.graph, post: async () => ok(answer) };
      expect(errorOf(await runWrite(graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }))).toMatchObject(
        Object.keys(answer).length === 0 ? { message: 'Graph moved the message but gave no new id.' } : { code: 'invalid_id' }
      );
    }
  });

  it('answers the first error when each write fails', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('PATCH /me/messages/m1', { type: 'api_error', status: 503, message: 'First.' });
    mailbox.failures.set('PATCH /me/messages/m2', { type: 'api_error', status: 503, message: 'Second.' });
    expect(errorOf(await runWrite(mailbox.graph, 'mark-mail-message', { messageId: 'm1,m2', read: 'true' })).message).toBe('First.');
  });

  it('reports a message that cannot be read again before its write, and writes the others', async () => {
    const mailbox = createMailStoreFake(SEED);
    const planned = await commands['mark-mail-message']?.plan?.(mailbox.graph, { messageId: 'm1,m2', read: 'true' });
    mailbox.failures.set('GET /me/messages/m1', { type: 'api_error', status: 503, message: 'Busy.' });
    const applied = planned?.ok === true ? await planned.value.apply() : undefined;
    expect(applied?.ok === true ? applied.value.value : applied).toMatchObject({ done: 1, failed: [{ target: '/me/messages/m1', reason: 'Busy.' }] });
  });

  it('passes over a message that moved to another folder with the same changeKey', async () => {
    const mailbox = createMailStoreFake(SEED);
    const planned = await commands['mark-mail-message']?.plan?.(mailbox.graph, { messageId: 'm1', read: 'true' });
    mailbox.messages.set('m1', { id: 'm1', parentFolderId: 'archive-id', changeKey: 'm1-ck', isRead: false });
    const applied = planned?.ok === true ? await planned.value.apply() : undefined;
    expect(applied?.ok === true ? applied.value.value : applied).toMatchObject({ done: 0, skipped: [{ target: '/me/messages/m1' }] });
  });
});

describe('the undo of a state field', () => {
  it('compares a due day as Graph gives it back, with seven digits of a second', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm1', status: 'flagged', due: '2026-11-20' }, { journal });
    mailbox.edit('m1', { flag: { flagStatus: 'flagged', dueDateTime: { dateTime: '2026-11-20T00:00:00.0000000', timeZone: 'UTC' } } });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
  });

  it('writes back the empty value of a field that Graph did not give', async () => {
    const mailbox = createMailStoreFake([{ id: 'm3', parentFolderId: 'inbox-id', changeKey: 'ck' }]);
    for (const [name, params, field, empty] of [
      ['set-mail-flag', { status: 'flagged' }, 'flag', { flagStatus: 'notFlagged' }],
      ['set-mail-categories', { categories: 'Red' }, 'categories', []],
      ['mark-mail-message', { read: 'true' }, 'isRead', false],
    ] as const) {
      const journal = createRunJournalFake();
      const run = await runWrite(mailbox.graph, name, { messageId: 'm3', ...params }, { journal });
      expect([name, run['runId']]).toEqual([name, RUN_ID]);
      expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
      expect(mailbox.messages.get('m3')?.[field]).toEqual(empty);
    }
  });

  it('refuses undo data that a person edited: an id that is not an id, or a field that no command changes', async () => {
    for (const data of [
      { messageId: 'm1/x', field: 'isRead', before: false, set: true },
      { messageId: 'm1', field: 'subject', before: 'a', set: 'b' },
    ]) {
      const mailbox = createMailStoreFake(SEED);
      const journal = createRunJournalFake();
      journal.records.set(RUN_ID, {
        version: 1,
        runId: RUN_ID,
        command: 'mark-mail-message',
        startedAt: '2026-11-18T09:05:07.000Z',
        changes: [{ action: 'mark-message', target: '/x', before: {}, after: {}, undo: { kind: 'restore-message-field', data } }],
      });
      expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual([
        'failed: The run journal names a message id that is not a Graph id, or a field that no command changes, so this undo does not run.',
      ]);
      expect(mailbox.calls).toEqual([]);
    }
  });

  it('reads a check answer it cannot read as a change', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'delete-mail-draft', { messageId: 'm1' }, { journal });
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }, { journal });
    const odd = answering(mailbox.graph, (path) => (path.includes('$select=parentFolderId,isDraft') ? ok('nothing') : undefined));
    expect(statusesOf(await undoWrite(odd, journal, true))).toEqual([expect.stringContaining('changed:')]);
  });
});

describe('the execute of a library caller', () => {
  it('makes the changes of the plan and answers them with the ids before and after', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answer = await commands['mark-mail-message']?.execute(mailbox.graph, { messageId: 'm1', read: 'true' });
    expect(answer).toEqual(
      ok({
        done: 1,
        skipped: [],
        failed: [],
        changes: [
          {
            action: 'mark-message',
            target: '/me/messages/m1',
            before: { messageId: 'm1', folderId: 'inbox-id', folder: 'Inbox', isRead: 'false' },
            after: { messageId: 'm1', isRead: 'true' },
          },
        ],
      })
    );
  });

  it('answers the refusal of the plan, and the error of the writes', async () => {
    const mailbox = createMailStoreFake(SEED);
    expect((await commands['mark-mail-message']?.execute(mailbox.graph, { read: 'true' }))?.ok).toBe(false);
    mailbox.failures.set('PATCH /me/messages/m1', { type: 'api_error', status: 503, message: 'Busy.' });
    const failed = await commands['mark-mail-message']?.execute(mailbox.graph, { messageId: 'm1', read: 'true' });
    expect(failed?.ok === false ? failed.error.message : failed).toBe('Busy.');
  });
});

describe('the id of a Graph answer', () => {
  it('is taken when it has the form of an id, and is a defect when it has not', () => {
    expect(graphId('drafts')).toBe(graphId('drafts'));
    expect(() => graphId('a/b')).toThrow('Graph answer: "a/b" is not a Graph id.');
  });
});
