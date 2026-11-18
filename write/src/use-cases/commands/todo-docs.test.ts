import { describe, expect, it } from 'bun:test';
import { commands } from './write-commands.ts';

/*
 * The docs of the To Do commands: what each undo does (and where it is
 * partial), the Graph path and the body that each write sends. The docs, the
 * help and the MCP tools show these words.
 */

const FIELDS =
  "{ title{T}: '{title}', body?: { content: '{body}', contentType: 'text' }, status?: '{status}', importance?: '{importance}', categories?: [{categories}], dueDateTime?: { dateTime: '{due}T00:00:00', timeZone: <time zone of the mailbox> }, startDateTime?: { dateTime: '{start}T00:00:00', timeZone: <time zone of the mailbox> }, reminderDateTime?: { dateTime: '{reminder}:00', timeZone: <time zone of the mailbox> }, isReminderOn?: true }";

const TASK = '/me/todo/lists/{list}/tasks/{task-id}';

describe('the docs of the To Do commands', () => {
  it('say what each undo does, where it writes, and what it sends', () => {
    const docs = Object.fromEntries(
      Object.entries(commands)
        .filter(([name]) => name.includes('-todo-'))
        .map(([name, command]) => [name, [command.meta.undo, command.meta.graphPathTemplate, command.meta.bodyTemplate]])
    );
    expect(docs).toEqual({
      'create-todo-task': ['deletes the task, only if it did not change after the run', '/me/todo/lists/{list}/tasks', FIELDS.replace('{T}', '')],
      'update-todo-task': ['writes the fields of before back, only if each field still has the value that the run set', TASK, FIELDS.replace('{T}', '?')],
      'complete-todo-task': ['gives the task its status of before, only if it is still completed', TASK, "{ status: 'completed' }"],
      'move-todo-task': [
        'moves the task back with a new copy (it gets a new id again), only if it did not change after the run. The dates of creation are not kept',
        '/me/todo/lists/{to}/tasks, then DELETE /me/todo/lists/{list}/tasks/{task-id}',
        '{ <the fields of the task> }, then one POST for each step, linked resource and attachment',
      ],
      'delete-todo-task': [
        'creates the task again from the copy in the run journal, with its steps, linked resources and attachments. It is a new task: it gets a new id, and the dates of creation and the date a step was checked are not kept',
        TASK,
        undefined,
      ],
      'add-todo-linked-resource': [
        'deletes the linked resource, only if it did not change after the run',
        `${TASK}/linkedResources`,
        "{ displayName: '{display-name}', applicationName: '{application-name}', webUrl?: '{web-url}', externalId?: '{external-id}' }",
      ],
      'apply-todo-ops': [
        'reverses each change of the run as the single commands do: it deletes an added task, writes the fields of before back, deletes the new steps, moves a task back, and creates a deleted task again (with a new id)',
        '/me/todo/lists/<list id>/tasks (and PATCH, DELETE on /me/todo/lists/<list id>/tasks/<task id>)',
        'for each op of {file}: the body of create-todo-task, update-todo-task, add-todo-task-step, move-todo-task or delete-todo-task',
      ],
      'add-todo-task-step': [
        'deletes the step, only if its text and its check mark did not change after the run',
        `${TASK}/checklistItems`,
        "{ displayName: '{title}', isChecked: {checked} }",
      ],
      'update-todo-task-step': [
        'gives the step its text and its check mark of before, only if they did not change after the run. The date the step was checked is not kept',
        `${TASK}/checklistItems/{step-id}`,
        "{ displayName?: '{title}', isChecked?: {checked} }",
      ],
      'delete-todo-task-step': [
        'adds the step again with its text and its check mark, at the end of the steps. It gets a new id, and the date it was checked is not kept',
        `${TASK}/checklistItems/{step-id}`,
        undefined,
      ],
      'add-todo-task-attachment': [
        'removes the attachment, only if it is still there with its name',
        `${TASK}/attachments (or attachments/createUploadSession for 3 MB or more)`,
        "{ '@odata.type': '#microsoft.graph.taskFileAttachment', name: '<the name of {file}>', contentBytes: '<the bytes of {file}, base64>' }",
      ],
      'remove-todo-task-attachment': [
        'attaches the file again from the bytes in the run journal. It is a new attachment with a new id',
        `${TASK}/attachments/{attachment-id}`,
        undefined,
      ],
      'create-todo-list': [
        'does not delete the list: no command of this CLI deletes a list. The undo refuses and says so; delete the list in To Do',
        '/me/todo/lists',
        "{ displayName: '{name}' }",
      ],
      'update-todo-list': ['gives the list its name of before, only if it still has the name that the run gave', '/me/todo/lists/{list}', "{ displayName: '{name}' }"],
    });
  });

  it('put the commands that delete or move on the destructive class, and the others on task-write', () => {
    const effects = Object.fromEntries(
      Object.entries(commands)
        .filter(([name]) => name.includes('-todo-'))
        .map(([name, command]) => [name, command.meta.effect])
    );
    expect(
      Object.entries(effects)
        .filter(([, effect]) => effect === 'delete')
        .map(([name]) => name)
        .toSorted((a, b) => a.localeCompare(b))
    ).toEqual(['apply-todo-ops', 'delete-todo-task', 'delete-todo-task-step', 'move-todo-task', 'remove-todo-task-attachment']);
    expect(Object.values(effects).filter((effect) => effect !== 'delete' && effect !== 'task-write')).toEqual([]);
  });
});
