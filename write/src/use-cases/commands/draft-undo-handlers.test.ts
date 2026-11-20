import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { undoRun } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { UndoReport } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { RunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailboxFake } from '../../test-helpers/mailbox-fake.ts';
import { DRAFT_CREATE_UNDO, DRAFT_UPDATE_UNDO } from './draft-runs.ts';
import { commands, undoHandlers } from './write-commands.ts';

/*
 * The undo handlers of the draft commands, on a journal record that a test
 * writes (the file on disk is outside data), and the docs that say what each
 * undo does.
 */

const RUN = '20261118T090507Z-0a1b2c3d';
const NOW = new Date('2026-11-18T09:05:07.000Z');

const undo = async (graph: WriteGraph, journal: RunJournalFake, apply: boolean): Promise<Result<UndoReport, unknown>> =>
  undoRun({ journal, graph, handlers: undoHandlers, now: () => NOW }, RUN, apply);

describe('the undo handlers of the drafts', () => {
  const recordWith = (kind: string, data: Record<string, unknown>): RunJournalFake => {
    const journal = createRunJournalFake();
    const change = { action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1' }, undo: { kind, data } };
    journal.records.set(RUN, { version: 1, runId: RUN, command: 'create-mail-draft', startedAt: NOW.toISOString(), changes: [change] });
    return journal;
  };

  it('refuses a message id from the journal that is not a Graph id, before any request', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-1', isDraft: true, changeKey: 'k' }]);
    for (const kind of ['delete-draft', 'restore-draft']) {
      for (const messageId of ['../mailFolders/inbox', 'draft-1/../../mailFolders/inbox']) {
        const result = await undo(mailbox.graph, recordWith(kind, { messageId, changeKey: 'k', fields: {} }), true);
        expect(result.ok ? result.value.changes.map((change) => [change.status, change.reason]) : []).toEqual([
          ['failed', 'The run journal names a message id that is not a Graph message id, so this undo does not run.'],
        ]);
      }
    }
    expect(mailbox.calls).toEqual([]);
  });

  it('refuses to apply an undo step whose data lost its shape', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-1', isDraft: true, changeKey: 'k' }]);
    for (const [kind, data] of [
      ['delete-draft', { changeKey: 'k' }],
      ['restore-draft', { messageId: 'draft-1', changeKey: 'k' }],
    ] as const) {
      const handler = undoHandlers[kind];
      if (handler === undefined) throw new Error(kind);
      expect(await handler.apply(mailbox.graph, data)).toEqual(
        err({ type: 'validation_error', message: 'The run journal names a message id that is not a Graph message id, so this undo does not run.' })
      );
    }
    expect(mailbox.calls).toEqual([]);
  });

  it('reports a read of the draft that fails for another reason than a missing draft', async () => {
    const mailbox = createMailboxFake();
    const down: WriteGraph = { ...mailbox.graph, get: async () => err({ type: 'api_error', status: 503, message: 'Service unavailable.' }) };
    const result = await undo(down, recordWith('delete-draft', { messageId: 'draft-1', changeKey: 'k' }), true);
    expect(result.ok ? result.value.changes.map((change) => [change.status, change.reason]) : []).toEqual([['failed', 'Service unavailable.']]);
  });

  it('passes over a draft whose version cannot be read, and a draft that is gone', async () => {
    const mailbox = createMailboxFake();
    const odd: WriteGraph = { ...mailbox.graph, get: async () => ok({ isDraft: true }) };
    const unreadable = await undo(odd, recordWith('delete-draft', { messageId: 'draft-1', changeKey: 'k' }), true);
    const gone = await undo(mailbox.graph, recordWith('delete-draft', { messageId: 'draft-1', changeKey: 'k' }), true);
    expect([unreadable, gone].map((result) => (result.ok ? result.value.changes.map((change) => change.status) : []))).toEqual([['changed'], ['missing']]);
  });
});

describe('the docs of the draft commands', () => {
  it('say what the undo of each one does', () => {
    const drafts = ['create-mail-draft', 'create-reply-draft', 'create-forward-draft', 'update-mail-draft'];
    expect(drafts.map((name) => commands[name]?.meta.undo)).toEqual([DRAFT_CREATE_UNDO, DRAFT_CREATE_UNDO, DRAFT_CREATE_UNDO, DRAFT_UPDATE_UNDO]);
    expect([DRAFT_CREATE_UNDO, DRAFT_UPDATE_UNDO]).toEqual([
      'moves the draft to Deleted Items (`DELETE /me/messages/{id}`), only if the message is still a draft and its changeKey did not change after the run',
      'writes back the fields that the update changed, as they were before the run, only if the message is still a draft and its changeKey did not change after the run',
    ]);
  });

  it('give a plan exactly to the commands whose docs say what the undo does, and to no PDF conversion', () => {
    const planned = Object.entries(commands).filter(([, command]) => command.plan !== undefined);
    expect(planned.map(([name]) => name).toSorted((a, b) => a.localeCompare(b))).toEqual([
      'add-mail-draft-attachment',
      'add-planner-checklist-item',
      'add-planner-task-file',
      'add-planner-task-link',
      'add-todo-linked-resource',
      'add-todo-task-attachment',
      'add-todo-task-step',
      'apply-todo-ops',
      'complete-todo-task',
      'create-calendar-draft',
      'create-forward-draft',
      'create-mail-draft',
      'create-mail-folder',
      'create-planner-bucket',
      'create-planner-task',
      'create-reply-draft',
      'create-todo-list',
      'create-todo-task',
      'delete-calendar-draft',
      'delete-mail-draft',
      'delete-planner-bucket',
      'delete-planner-task',
      'delete-todo-task',
      'delete-todo-task-step',
      'mark-mail-message',
      'move-mail-message',
      'move-mail-thread',
      'move-todo-task',
      'remove-mail-draft-attachment',
      'remove-planner-checklist-item',
      'remove-planner-task-link',
      'remove-todo-task-attachment',
      'set-mail-categories',
      'set-mail-flag',
      'update-calendar-draft',
      'update-mail-draft',
      'update-planner-bucket',
      'update-planner-checklist-item',
      'update-planner-plan',
      'update-planner-task',
      'update-planner-task-details',
      'update-todo-list',
      'update-todo-task',
      'update-todo-task-step',
    ]);
    expect(Object.values(commands).filter((command) => (command.plan === undefined) === (command.meta.undo !== undefined))).toEqual([]);
  });
});
