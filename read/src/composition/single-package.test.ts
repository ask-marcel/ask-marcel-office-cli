import { describe, expect, it } from 'bun:test';
import { commands } from '../use-cases/commands/index.ts';
import { LIFECYCLE_ENTRIES } from '../use-cases/commands/lifecycle-entries.ts';
import { commandPackage } from './single-package.ts';

// The names are pinned as literals, not read from package.json: the Stryker
// sandbox copies no *.json, so a test that imports package.json fails its dry run.
describe('the single package the shells are wired with', () => {
  it('ships the whole registry and the lifecycle entries', () => {
    expect(commandPackage.registry).toBe(commands);
    expect(commandPackage.lifecycle).toBe(LIFECYCLE_ENTRIES);
  });

  it('keeps the bin and the package name it has always published', () => {
    expect(commandPackage.binName).toBe('ask-marcel-office');
    expect(commandPackage.packageName).toBe('ask-marcel-office-cli');
  });
});
