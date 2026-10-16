import { describe, expect, it } from 'bun:test';
import { COMMAND_DIRECTORY, commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { commands } from './index.ts';
import { LIFECYCLE_ENTRIES } from './lifecycle-entries.ts';

// The directory is static so core can read it without the registry; this suite
// is what keeps it true to the registry the single package ships.
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

  it('names no command that the registry, the lifecycle entries and the token helper do not hold', () => {
    const shipped = new Set([...Object.keys(commands), ...LIFECYCLE_ENTRIES.map((entry) => entry.name), 'token']);
    expect([...COMMAND_DIRECTORY.keys()].filter((name) => !shipped.has(name))).toEqual([]);
  });

  it('leaves out the commands every package has, so no package sends a caller elsewhere for them', () => {
    expect(['docs', 'help', 'help-json', 'mcp'].map((name) => commandRole(name))).toEqual([undefined, undefined, undefined, undefined]);
  });

  it('knows no command by a name that is not a command, such as a prototype key', () => {
    expect([commandRole('toString'), commandRole('constructor'), commandRole('list-widgets')]).toEqual([undefined, undefined, undefined]);
  });
});
