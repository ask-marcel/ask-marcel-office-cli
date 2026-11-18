import { describe, expect, it } from 'bun:test';
import { garbageSweep, runWithFault, undoWithFault, unseenFailures } from '../../test-helpers/todo-faults.ts';
import { SCENARIOS } from '../../test-helpers/todo-scenarios.ts';

/*
 * The fault sweeps of the To Do commands. Each Graph call of a run and of its
 * undo fails in turn: the outcome shows that failure, so no command goes on
 * after a failed read or write. Each call gets an answer of a wrong shape in
 * turn: the outcome names the item that Graph gave in that shape, or nothing
 * when the CLI does not need the answer (a delete, a step it added in a copy).
 */

const UNREADABLE: Readonly<Record<string, { readonly run: ReadonlyArray<string>; readonly undo: ReadonlyArray<string> }>> = {
  'create-todo-task': { run: ['the To Do lists', 'the time zone of the mailbox', 'the task it created'], undo: ['a To Do task', '-'] },
  'update-todo-task': { run: ['the To Do lists', 'a To Do task', 'the time zone of the mailbox', 'a To Do task', '-'], undo: ['a To Do task', '-'] },
  'complete-todo-task': { run: ['the To Do lists', 'a To Do task', 'a To Do task', '-'], undo: ['a To Do task', '-'] },
  'move-todo-task': {
    run: [
      'the To Do lists',
      'a To Do task',
      'the To Do lists',
      'the steps of a To Do task',
      'the linked resources of a To Do task',
      'the attachments of a To Do task',
      'a To Do task',
      'the steps of a To Do task',
      'the linked resources of a To Do task',
      'the attachments of a To Do task',
      'an attachment of a To Do task',
      'an attachment of a To Do task',
      'the task it created',
      '-',
      '-',
      '-',
      '-',
      '-',
      'a To Do task',
      '-',
    ],
    undo: [
      '-',
      'a To Do task',
      'a To Do task',
      'the steps of a To Do task',
      'the linked resources of a To Do task',
      'the attachments of a To Do task',
      'an attachment of a To Do task',
      'an attachment of a To Do task',
      'the task it created',
      '-',
      '-',
      '-',
      '-',
      '-',
      'a To Do task',
      '-',
    ],
  },
  'delete-todo-task': {
    run: [
      'the To Do lists',
      'a To Do task',
      'the steps of a To Do task',
      'the linked resources of a To Do task',
      'the attachments of a To Do task',
      'a To Do task',
      'the steps of a To Do task',
      'the linked resources of a To Do task',
      'the attachments of a To Do task',
      'an attachment of a To Do task',
      'an attachment of a To Do task',
      '-',
    ],
    undo: ['-', 'the task it created', '-', '-', '-', '-', '-', 'a To Do task'],
  },
  'add-todo-linked-resource': { run: ['the To Do lists', 'a To Do task', 'the linked resource it created'], undo: ['-', '-'] },
  'add-todo-task-step': { run: ['the To Do lists', 'a To Do task', 'a step of a To Do task'], undo: ['a To Do task', 'a step of a To Do task', '-'] },
  'update-todo-task-step': {
    run: ['the To Do lists', 'a To Do task', 'a step of a To Do task', 'a step of a To Do task', 'a step of a To Do task'],
    undo: ['a To Do task', 'a step of a To Do task', '-'],
  },
  'delete-todo-task-step': { run: ['the To Do lists', 'a To Do task', 'a step of a To Do task', 'a step of a To Do task', '-'], undo: ['a To Do task', '-'] },
  'add-todo-task-attachment': { run: ['the To Do lists', 'a To Do task', '-'], undo: ['a To Do task', 'the attachments of a To Do task', '-'] },
  'remove-todo-task-attachment': {
    run: ['the To Do lists', 'a To Do task', 'the attachments of a To Do task', 'the attachments of a To Do task', 'the bytes of the attachment', '-'],
    undo: ['a To Do task', '-'],
  },
  'create-todo-list': { run: ['the list it created'], undo: [] },
  'update-todo-list': { run: ['the To Do lists', 'a To Do list', 'the list it changed'], undo: ['a To Do list', '-'] },
  'apply-todo-ops': {
    run: ['the time zone of the mailbox', 'the To Do lists', 'a To Do task', 'the task it created', 'a To Do task', '-', 'a step of a To Do task'],
    undo: ['a To Do task', 'a step of a To Do task', '-', 'a To Do task', '-', 'a To Do task', '-'],
  },
};

describe('the fault sweeps of the To Do commands', () => {
  for (const [name, scenario] of Object.entries(SCENARIOS)) {
    it(`${name}: shows each failed call of the run and of its undo`, async () => {
      expect(await unseenFailures(async (at, fault) => runWithFault(scenario, at, fault))).toEqual([]);
      const undo = await unseenFailures(async (at, fault) => undoWithFault(scenario, at, fault));
      expect(undo).toEqual(name === 'create-todo-list' ? [-1] : []);
    });

    it(`${name}: names the item of each answer that it cannot read`, async () => {
      expect({
        run: await garbageSweep(async (at, fault) => runWithFault(scenario, at, fault)),
        undo: await garbageSweep(async (at, fault) => undoWithFault(scenario, at, fault)),
      }).toEqual(UNREADABLE[name]);
    });
  }
});
