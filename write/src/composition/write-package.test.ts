import { describe, expect, it } from 'bun:test';
import { fillNames } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { commands } from '../use-cases/commands/write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from '../use-cases/commands/write-lifecycle-entries.ts';
import { commandPackage } from './write-package.ts';
import { WRITE_PACKAGE_NAMES } from './write-package-names.ts';

// The write package `@ask-marcel/office-write` as the core shells see it. The
// names are pinned as literals, not read from package.json: the Stryker sandbox
// copies no *.json, so a test that imports package.json fails its dry run.
describe('the write package the shells are wired with', () => {
  it('ships the write registry and the write lifecycle entries, and no read command', () => {
    expect(commandPackage.registry).toBe(commands);
    expect(commandPackage.lifecycle).toBe(WRITE_LIFECYCLE_ENTRIES);
    expect(Object.values(commandPackage.registry).filter((command) => command.meta.effect === 'read')).toEqual([]);
  });

  it('names the write bin as its own bin, the bins of auth and read as theirs, and its npm name', () => {
    expect([commandPackage.binName, commandPackage.bins, commandPackage.packageName]).toEqual([
      'ask-marcel-office-write',
      { auth: 'ask-marcel-office-auth', read: 'ask-marcel-office-read', write: 'ask-marcel-office-write' },
      '@ask-marcel/office-write',
    ]);
  });

  it('gives a write command of its own registry as the docs example', () => {
    expect(commandPackage.exampleCommand).toBe('create-mail-draft');
    expect(Object.hasOwn(commandPackage.registry, commandPackage.exampleCommand)).toBe(true);
  });

  it('says in its own words what it writes, that it sends nothing, and where the reads and the sign-in are', () => {
    const { description } = commandPackage;
    expect(description).toContain('mail drafts');
    expect(description).toContain('It sends no mail');
    expect(description).toContain('@ask-marcel/office-read');
    expect(description).toContain('@ask-marcel/office-auth');
  });

  it('names packages in its description, never a bin, because the shell does not fill it', () => {
    const { description } = commandPackage;
    expect(description).not.toMatch(/ask-marcel-office-|\{[a-z-]*bin\}/);
    // cli-write.test.ts refuses these words in the top-level description.
    expect(description).not.toContain('keeps no file');
    expect(description).not.toContain('write files');
  });

  it('fills a read command reference with the read bin and its own with the write bin', () => {
    expect(fillNames('Run `{read-bin} find-mail-drafts`, then `{bin} docs create-mail-draft`.', WRITE_PACKAGE_NAMES)).toBe(
      'Run `ask-marcel-office-read find-mail-drafts`, then `ask-marcel-office-write docs create-mail-draft`.'
    );
  });
});
