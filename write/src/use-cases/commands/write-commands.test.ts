import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import { commands } from './write-commands.ts';

// The registry of the write package `@ask-marcel/office-write` (package split,
// phase 5): its 46 commands, each wrapped once so that an unknown parameter
// is refused and the command runs on the write graph alone.

const WRITE_COMMANDS = [
  'add-mail-draft-attachment',
  'add-planner-checklist-item',
  'add-planner-task-file',
  'add-planner-task-link',
  'add-todo-linked-resource',
  'add-todo-task-attachment',
  'add-todo-task-step',
  'apply-todo-ops',
  'complete-todo-task',
  'convert-calendar-event-attachment-to-pdf',
  'convert-group-post-attachment-to-pdf',
  'convert-mail-attachment-to-pdf',
  'create-calendar-draft',
  'create-forward-draft',
  'create-mail-draft',
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
];

// A write graph that records each member that a command calls.
const recordingWriteGraph = (called: string[]): WriteGraph => {
  const answer = (member: string): (() => Promise<ReturnType<typeof ok<object>>>) => {
    const recorded = async (): Promise<ReturnType<typeof ok<object>>> => {
      called.push(member);
      return ok({});
    };
    return recorded;
  };
  return fakeWriteGraph(Object.fromEntries(['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put'].map((member) => [member, answer(member)])));
};

describe('the registry of the write package', () => {
  it('holds the 46 write commands, and no read command', () => {
    expect(Object.keys(commands).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_COMMANDS);
    expect(Object.values(commands).filter((command) => command.meta.effect === 'read')).toEqual([]);
  });

  for (const name of WRITE_COMMANDS) {
    it(`refuses an unknown parameter of \`${name}\` before it calls Graph`, async () => {
      const called: string[] = [];
      const result = await commands[name]?.execute(recordingWriteGraph(called), { draftBodyy: 'x' });
      expect(result?.ok === false && result.error.code).toBe('unknown_parameter');
      expect(called).toEqual([]);
    });
  }

  it('runs a draft on the write graph alone', async () => {
    const called: string[] = [];
    const result = await commands['create-mail-draft']?.execute(recordingWriteGraph(called), { subject: 'Agenda', bodyContent: 'Hi', toRecipients: 'robin.chen@contoso.com' });
    expect(result?.ok).toBe(true);
    expect(called).toEqual(['post']);
  });
});
