import { describe, expect, it } from 'bun:test';
import type { AttachmentFile } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { changingAt, recordingUploads, withItems } from '../../test-helpers/todo-scenarios.ts';
import { createSeededTodoStore, TIME_ZONE } from '../../test-helpers/todo-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { undoHandlers } from './write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from './write-lifecycle-entries.ts';

/*
 * The rules of the steps, the attachments, the lists, the undo of a move and
 * the ops file that the main scenarios do not show: a change between the plan
 * and the write, an answer with no id, and the exact words of the refusals.
 */

const TASK = { list: 'Tasks', taskId: 'task-a' };

describe('the steps of a task', () => {
  it('pass over a step whose text or check mark changed between the plan and the write', async () => {
    for (const change of [{ displayName: 'Other' }, { isChecked: true }]) {
      const todo = createSeededTodoStore();
      const step = todo.steps.get('step-a') ?? { id: 'step-a' };
      const graph = changingAt(todo, '/checklistItems/step-a', 2, () => todo.steps.set('step-a', { ...step, ...change }));
      const run = await runWrite(graph, 'delete-todo-task-step', { ...TASK, stepId: 'step-a' });
      expect(run).toMatchObject({ done: 0, skipped: [{ reason: expect.stringMatching(/^The item changed after the plan/) }] });
    }
  });

  it('keep the check mark when only the text changes', async () => {
    const todo = createSeededTodoStore();
    todo.steps.set('step-a', { ...todo.steps.get('step-a'), id: 'step-a', isChecked: true });
    await runWrite(todo.graph, 'update-todo-task-step', { ...TASK, stepId: 'step-a', title: 'New text' });
    expect(todo.steps.get('step-a')).toMatchObject({ displayName: 'New text', isChecked: true });
  });

  it('refuse a step that the task does not have', async () => {
    const result = await runWrite(createSeededTodoStore().graph, 'delete-todo-task-step', { ...TASK, stepId: 'step-z' });
    expect(result['error']).toMatchObject({ message: 'The specified object was not found in the store.' });
  });

  it('say no new id when Graph gives the step back with no id', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'delete-todo-task-step', { ...TASK, stepId: 'step-a' }, { journal });
    const graph: WriteGraph = { ...todo.graph, post: async () => ok({ displayName: 'Read the draft' }) };
    const undo = await undoWrite(graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', undefined]]);
  });
});

describe('the attachments of a task', () => {
  const files = createLocalFilesFake({ 'a.txt': 'abc', 'c.txt': 'abc', 'big.bin': 'x'.repeat(25_000_001) });

  it('remove the attachment that the id names, of two', async () => {
    const todo = createSeededTodoStore();
    await withItems(todo);
    const run = await runWrite(todo.graph, 'remove-todo-task-attachment', { ...TASK, attachmentId: 'file-x' });
    expect(run).toMatchObject({ done: 1, changes: [{ before: { file: 'b.bin' } }] });
    expect([...todo.files.values()].map((file) => file['name'])).toEqual(['a.txt']);
  });

  it('pass over an attachment whose size or name changed between the plan and the write', async () => {
    for (const change of [{ size: 4 }, { name: 'c.bin' }]) {
      const todo = createSeededTodoStore();
      await withItems(todo);
      const file = todo.files.get('file-x') ?? { id: 'file-x' };
      const graph = changingAt(todo, '/attachments', 2, () => todo.files.set('file-x', { ...file, ...change }));
      expect(await runWrite(graph, 'remove-todo-task-attachment', { ...TASK, attachmentId: 'file-x' })).toMatchObject({
        done: 0,
        skipped: [{ reason: expect.stringMatching(/^The item changed/) }],
      });
    }
  });

  it('keep the content type of a removed file for the undo, and give none when Graph gives none', async () => {
    const todo = createSeededTodoStore();
    await withItems(todo);
    const id = [...todo.files.values()].find((file) => file['name'] === 'a.txt')?.id ?? '';
    for (const attachmentId of [id, 'file-x']) {
      const journal = createRunJournalFake();
      await runWrite(todo.graph, 'remove-todo-task-attachment', { ...TASK, attachmentId }, { journal });
      const uploads: AttachmentFile[] = [];
      await undoWrite(recordingUploads(todo, uploads), journal, true);
      expect(uploads.map((file) => file.contentType ?? 'none')).toEqual([attachmentId === 'file-x' ? 'none' : 'text/plain']);
    }
  });

  it('undo the add of an attachment beside another one, and give no new id when the upload gives none', async () => {
    const todo = createSeededTodoStore();
    await withItems(todo);
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'add-todo-task-attachment', { ...TASK, file: 'c.txt' }, { journal, files });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual(['will-undo']);
    const second = createRunJournalFake();
    await runWrite(todo.graph, 'remove-todo-task-attachment', { ...TASK, attachmentId: 'file-x' }, { journal: second });
    const noId: WriteGraph = { ...todo.graph, uploadAttachment: async (_target, file) => ok({ name: file.name, size: file.bytes.byteLength }) };
    const undo = await undoWrite(noId, second, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', undefined]]);
  });

  it('give an empty id when the upload gives none', async () => {
    const todo = createSeededTodoStore();
    const noId: WriteGraph = { ...todo.graph, uploadAttachment: async (_target, file) => ok({ name: file.name, size: file.bytes.byteLength }) };
    const run = await runWrite(noId, 'add-todo-task-attachment', { ...TASK, file: 'a.txt' }, { files });
    expect(run['changes']).toEqual([expect.objectContaining({ after: { listId: 'list-a', taskId: 'task-a', attachmentId: '' } })]);
  });

  it('refuse a file above 25 MB, and a task that is not there', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'add-todo-task-attachment', { ...TASK, file: 'big.bin' }, { files, dryRun: true }))['error']).toMatchObject({
      code: 'local_file_too_large',
    });
    expect((await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-z', file: 'a.txt' }, { files }))['error']).toMatchObject({
      message: 'The specified object was not found in the store.',
    });
  });
});

describe('the lists', () => {
  it('pass over a list whose name changed between the plan and the write', async () => {
    const todo = createSeededTodoStore();
    const graph = changingAt(todo, '/me/todo/lists/list-b', 1, () => todo.lists.set('list-b', { id: 'list-b', displayName: 'By hand' }));
    expect(await runWrite(graph, 'update-todo-list', { list: 'Later', name: 'Someday' })).toMatchObject({
      done: 0,
      skipped: [{ reason: expect.stringMatching(/^The item changed/) }],
    });
  });

  it('show no change for a list that has the name already', async () => {
    const dry = await runWrite(createSeededTodoStore().graph, 'update-todo-list', { list: 'Later', name: 'Later' }, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'update-todo-list',
      changes: [],
      skipped: [{ target: '/me/todo/lists/list-b', before: { listId: 'list-b', list: 'Later' }, reason: 'The list has this name already.' }],
    });
  });

  it('refuse the undo of a new list with its code', async () => {
    expect(await undoHandlers['todo-list-created']?.check(createSeededTodoStore().graph, { listId: 'list-a' })).toEqual({
      ok: false,
      error: { type: 'validation_error', code: 'no_list_delete', message: expect.stringMatching(/^No command of this CLI deletes a To Do list/) },
    });
  });
});

describe('the undo of a move', () => {
  it('gives the error of Graph when the moved task is gone at the apply', async () => {
    const result = await undoHandlers['move-todo-task-back']?.apply(createSeededTodoStore().graph, { listId: 'list-b', taskId: 'task-z', stamp: 'x', backTo: 'list-a' });
    expect(result).toMatchObject({ ok: false, error: { message: 'The specified object was not found in the store.' } });
  });
});

describe('the ops file', () => {
  const opsFile = (ops: unknown): ReturnType<typeof createLocalFilesFake> => createLocalFilesFake({ 'ops.json': JSON.stringify({ ops }) });

  it('refuses with the code of the ops file', async () => {
    const run = await runWrite(createSeededTodoStore().graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile([{ op: 'update', list: 'Tasks', id: 'task-a' }]) });
    expect(run['error']).toEqual({ source: 'validation', code: 'invalid_ops_file', message: '--file: op 1 needs "set", "addChecklist", or both.' });
    const bad = await runWrite(createSeededTodoStore().graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile('x') });
    expect(bad['error']).toEqual({ source: 'validation', code: 'invalid_ops_file', message: '--file: ops must be of the type array.' });
  });

  it('takes two adds, and an update with new steps only, with no PATCH', async () => {
    const todo = createSeededTodoStore();
    const files = opsFile([
      { op: 'add', list: 'Tasks', title: 'A' },
      { op: 'add', list: 'Tasks', title: 'B' },
      { op: 'update', list: 'Tasks', id: 'task-a', addChecklist: ['Print it'] },
    ]);
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files });
    expect(run).toMatchObject({ done: 3 });
    expect((run['changes'] as ReadonlyArray<unknown>)[2]).toMatchObject({ action: 'add-step', before: { listId: 'list-a', taskId: 'task-a' } });
    expect(todo.calls.filter((call) => call.startsWith('PATCH'))).toEqual([]);
    expect([...todo.steps.values()].map((step) => [step['displayName'], step['isChecked']])).toEqual([
      ['Read the draft', false],
      ['Print it', false],
    ]);
  });

  it('sets a due day of an update in the time zone of the mailbox', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile([{ op: 'update', list: 'Tasks', id: 'task-a', set: { due: '2026-12-01' } }]) });
    expect(todo.tasks.get('task-a')?.['dueDateTime']).toEqual({ dateTime: '2026-12-01T00:00:00', timeZone: TIME_ZONE });
  });

  it('refuses a move or a delete of a task that is not there', async () => {
    for (const op of [
      { op: 'move', list: 'Tasks', id: 'task-z', to: 'Later' },
      { op: 'delete', list: 'Tasks', id: 'task-z' },
    ]) {
      const run = await runWrite(createSeededTodoStore().graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile([op]) });
      expect(run['error']).toMatchObject({ message: 'The specified object was not found in the store.' });
    }
  });

  it('gives the error of the time zone of the mailbox', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('GET /me/mailboxSettings/timeZone', { type: 'api_error', status: 403, message: 'No scope.' });
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile([{ op: 'add', list: 'Tasks', title: 'A', due: '2026-12-01' }]) });
    expect(run['error']).toMatchObject({ message: 'No scope.' });
  });

  it('records the changes of an update whose new step failed, with the error', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('POST /me/todo/lists/list-a/tasks/task-a/checklistItems', { type: 'api_error', status: 500, message: 'Step failed.' });
    const files = opsFile([{ op: 'update', list: 'Tasks', id: 'task-a', set: { title: 'B' }, addChecklist: ['x'] }]);
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files });
    expect(run).toMatchObject({ done: 0, failed: [{ reason: 'Step failed.' }], runId: RUN_ID });
  });
});

describe('the undo lifecycle command', () => {
  it('describes its option and its run id', () => {
    const undo = WRITE_LIFECYCLE_ENTRIES.find((entry) => entry.name === 'undo');
    expect([undo?.options[0]?.description, undo?.positionalArguments?.[0]?.description]).toEqual([
      'Do the undo. Without it, the command only shows what the undo would do, and changes nothing.',
      'The runId field of the output of the write to reverse, for example 20261118T090507Z-0a1b2c3d.',
    ]);
  });
});
