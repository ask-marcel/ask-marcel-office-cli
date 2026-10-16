import { describe, expect, it } from 'bun:test';
import type { PackageRole } from './command-directory.ts';
import { COMMAND_DIRECTORY, PACKAGE_NAMES, PACKAGE_ROLES, commandRole } from './command-directory.ts';

/*
 * The shape of the command-to-package directory (core, package split phase 3).
 * Core reads it with no registry, so these cases hold without one: each entry
 * is a command name, once, in a sorted list per package. The read package's
 * command-directory.test.ts keeps it true to the registry the package ships.
 */

const COMMAND_NAME = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;

const namesOf = (role: PackageRole): ReadonlyArray<string> => [...COMMAND_DIRECTORY].filter(([, placed]) => placed === role).map(([name]) => name);

describe('the command-to-package directory', () => {
  it('holds 4 auth commands, 204 read commands and 7 write commands', () => {
    expect(PACKAGE_ROLES.map((role) => [role, namesOf(role).length])).toEqual([
      ['auth', 4],
      ['read', 204],
      ['write', 7],
    ]);
  });

  it('names each command in kebab case, once', () => {
    const names = [...COMMAND_DIRECTORY.keys()];
    expect(names.filter((name) => !COMMAND_NAME.test(name))).toEqual([]);
    expect(names).toHaveLength(215);
  });

  it('lists the commands of each package in alphabetical order, so an entry has one place to go', () => {
    for (const role of PACKAGE_ROLES) expect(namesOf(role)).toEqual(namesOf(role).toSorted((a, b) => a.localeCompare(b)));
  });

  it('places the sign-in commands and the token helper in auth, and a draft and a conversion in write', () => {
    expect(namesOf('auth')).toEqual(['login', 'logout', 'status', 'token']);
    expect(['create-mail-draft', 'convert-mail-attachment-to-pdf', 'list-drives'].map((name) => commandRole(name))).toEqual(['write', 'write', 'read']);
  });

  it('names each package by its npm name', () => {
    expect(PACKAGE_NAMES).toEqual({ auth: '@ask-marcel/office-auth', read: '@ask-marcel/office-read', write: '@ask-marcel/office-write' });
  });
});
