import { describe, expect, it } from 'bun:test';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { describeWrites, WRITE_EFFECTS } from './command-effect.ts';

/*
 * The write classes of the mail, To Do and Planner commands to come: a state
 * change of a mail message, and a write to a To Do or Planner item. Neither
 * deletes an item or sends mail, so neither is destructive; a delete gets its
 * own destructive class with its first command.
 */

describe('the write classes of mail state and tasks', () => {
  it('describe what a command of each class does to the tenant', () => {
    const registry = {
      'set-widget-flag': fakeCommand({ effect: 'mail-state' }),
      'create-widget-task': fakeCommand({ effect: 'task-write' }),
    };
    expect(describeWrites(registry)).toBe(
      'set-widget-flag: change the state of a mail message: mark it read or unread, flag it, set its categories, or move it to another folder (nothing is sent). create-widget-task: create or change a To Do list, a To Do or Planner task, a Planner bucket or the names of a Planner plan, or a part of a task (a step, a checklist item, an attachment or a link, which a change can also remove). The file of a Planner task goes into the folder "Planner attachments" of your OneDrive'
    );
  });

  it('are not destructive', () => {
    expect([WRITE_EFFECTS['mail-state'].destructive, WRITE_EFFECTS['task-write'].destructive]).toEqual([false, false]);
  });
});
