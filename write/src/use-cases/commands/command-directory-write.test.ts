import { describe, expect, it } from 'bun:test';
import { COMMAND_DIRECTORY, commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { commands } from './write-commands.ts';

// The write half of the directory check (package split, phase 5). The directory
// is static, so core reads it without a registry; this suite keeps its write
// entries true to the registry that the write package ships.
describe('the write entries of the command-to-package directory', () => {
  it('places each command of the write registry in write', () => {
    expect(Object.keys(commands).filter((name) => commandRole(name) !== 'write')).toEqual([]);
  });

  it('names no write command that the write registry does not hold', () => {
    const writes = [...COMMAND_DIRECTORY].filter(([, role]) => role === 'write').map(([name]) => name);
    expect(writes.filter((name) => !Object.hasOwn(commands, name))).toEqual([]);
    expect(writes).toHaveLength(Object.keys(commands).length);
  });
});
