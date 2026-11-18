import type { AttachmentFile } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { Scenario } from './todo-faults.ts';
import type { TodoStoreFake } from './todo-store-fake.ts';

/**
 * One scenario for each To Do command, for the fault sweeps of the faults helper.
 * The task of the seed gets a checked step, a linked resource with all its
 * fields, and two attachments, one with no content type, so a copy reads
 * each kind of item.
 */
export const withItems = async (todo: TodoStoreFake): Promise<void> => {
  await todo.graph.post('/me/todo/lists/list-a/tasks/task-a/checklistItems', { displayName: 'Done step', isChecked: true });
  await todo.graph.post('/me/todo/lists/list-a/tasks/task-a/linkedResources', {
    displayName: 'Spec',
    applicationName: 'Browser',
    webUrl: 'https://example.com/spec',
    externalId: 'spec-1',
  });
  await todo.graph.uploadAttachment(
    { kind: 'todo-task', listId: 'list-a', taskId: 'task-a' },
    { name: 'a.txt', bytes: new TextEncoder().encode('hello'), contentType: 'text/plain' }
  );
  todo.files.set('file-x', { id: 'file-x', taskId: 'task-a', name: 'b.bin', size: 3, contentBytes: 'YWJj' });
};

const TASK = { list: 'Tasks', taskId: 'task-a' };

export const OPS = JSON.stringify({
  ops: [
    { op: 'add', list: 'Tasks', title: 'New', due: '2026-11-20' },
    { op: 'update', list: 'Tasks', id: 'task-a', set: { importance: 'high' }, addChecklist: ['x'] },
  ],
});

export const SCENARIOS: Readonly<Record<string, Scenario>> = {
  'create-todo-task': { command: 'create-todo-task', params: { list: 'Tasks', title: 'New', due: '2026-11-20' } },
  'update-todo-task': { command: 'update-todo-task', params: { ...TASK, due: '2026-11-21', title: 'B' } },
  'complete-todo-task': { command: 'complete-todo-task', params: TASK },
  'move-todo-task': { command: 'move-todo-task', params: { ...TASK, to: 'Later' }, prepare: withItems },
  'delete-todo-task': { command: 'delete-todo-task', params: TASK, prepare: withItems },
  'add-todo-linked-resource': { command: 'add-todo-linked-resource', params: { ...TASK, displayName: 'x', applicationName: 'y' } },
  'add-todo-task-step': { command: 'add-todo-task-step', params: { ...TASK, title: 'x' } },
  'update-todo-task-step': { command: 'update-todo-task-step', params: { ...TASK, stepId: 'step-a', checked: 'true' } },
  'delete-todo-task-step': { command: 'delete-todo-task-step', params: { ...TASK, stepId: 'step-a' } },
  'add-todo-task-attachment': { command: 'add-todo-task-attachment', params: { ...TASK, file: 'a.txt' }, files: createLocalFilesFake({ 'a.txt': 'abc' }) },
  'remove-todo-task-attachment': { command: 'remove-todo-task-attachment', params: { ...TASK, attachmentId: 'file-x' }, prepare: withItems },
  'create-todo-list': { command: 'create-todo-list', params: { name: 'Errands' } },
  'update-todo-list': { command: 'update-todo-list', params: { list: 'Later', name: 'Someday' } },
  'apply-todo-ops': { command: 'apply-todo-ops', params: { file: 'ops.json' }, files: createLocalFilesFake({ 'ops.json': OPS }) },
};

// The graph of the store that keeps the file of each upload.
export const recordingUploads = (todo: TodoStoreFake, files: AttachmentFile[]): WriteGraph => ({
  ...todo.graph,
  uploadAttachment: async (target, file) => {
    files.push(file);
    return todo.graph.uploadAttachment(target, file);
  },
});

// The graph of the store that runs `change` before the read number `at` (from 1) of a path that ends with `suffix`.
export const changingAt = (todo: TodoStoreFake, suffix: string, at: number, change: () => void): WriteGraph => {
  let seen = 0;
  return {
    ...todo.graph,
    get: async (path, headers) => {
      if (path.endsWith(suffix) && ++seen === at) change();
      return todo.graph.get(path, headers);
    },
  };
};

export const tasksIn = (todo: TodoStoreFake, listId: string): ReadonlyArray<Record<string, unknown>> => [...todo.tasks.values()].filter((task) => task['listId'] === listId);
