import { describe, expect, it } from 'bun:test';
import { COMMAND_DIRECTORY, commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { commands } from '../../composition/read-commands.ts';
import { LIFECYCLE_ENTRIES } from './lifecycle-entries.ts';

// The directory is static so core can read it without the registry; this suite
// is what keeps it true to the registry the read package ships.
describe('the command-to-package directory', () => {
  it('places every registry command in the package its effect chooses: read for a read, write for a write', () => {
    const misplaced = Object.entries(commands)
      .map(([name, command]) => ({ name, expected: command.meta.effect === 'read' ? 'read' : 'write', placed: commandRole(name) }))
      .filter(({ expected, placed }) => expected !== placed);
    expect(misplaced).toEqual([]);
  });

  it('places the sign-in commands and the token helper in auth', () => {
    expect(['login', 'logout', 'status', 'token'].map((name) => commandRole(name))).toEqual(['auth', 'auth', 'auth', 'auth']);
  });

  it('names no command that the registry, the lifecycle entries, the auth package and the write package do not hold', () => {
    const shipped = new Set([
      ...Object.keys(commands),
      ...LIFECYCLE_ENTRIES.map((entry) => entry.name),
      'login',
      'logout',
      'status',
      'token',
      'add-mail-draft-attachment',
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
      'create-reply-draft',
      'create-todo-list',
      'create-todo-task',
      'delete-calendar-draft',
      'delete-mail-draft',
      'delete-todo-task',
      'delete-todo-task-step',
      'mark-mail-message',
      'move-mail-message',
      'move-mail-thread',
      'move-todo-task',
      'remove-mail-draft-attachment',
      'remove-todo-task-attachment',
      'set-mail-categories',
      'set-mail-flag',
      'update-calendar-draft',
      'update-mail-draft',
      'update-todo-list',
      'update-todo-task',
      'update-todo-task-step',
    ]);
    expect([...COMMAND_DIRECTORY.keys()].filter((name) => !shipped.has(name))).toEqual([]);
  });

  it('leaves out the commands every package has, so no package sends a caller elsewhere for them', () => {
    expect(['docs', 'help', 'help-json', 'mcp'].map((name) => commandRole(name))).toEqual([undefined, undefined, undefined, undefined]);
  });

  it('knows no command by a name that is not a command, such as a prototype key', () => {
    expect([commandRole('toString'), commandRole('constructor'), commandRole('list-widgets')]).toEqual([undefined, undefined, undefined]);
  });
});
