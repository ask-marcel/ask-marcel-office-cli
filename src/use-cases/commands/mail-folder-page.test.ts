import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const pathOf = async (name: string, params: Record<string, string>): Promise<string> => {
  let seen = '';
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  await command.execute(
    fakeGraphClient({
      get: async (path: string) => {
        seen = path;
        return ok({ value: [] });
      },
    }),
    params
  );
  return seen;
};

// Graph pages mail folders ten at a time, so an eleventh top-level folder went
// unseen by a caller that did not follow the cursor (a daily-brief run, 2026-09-24).
describe('a mail-folder listing', () => {
  it('asks for a hundred folders a page unless --top says otherwise', async () => {
    expect(await pathOf('list-mail-folders', {})).toBe('/me/mailFolders?$top=100');
    expect(await pathOf('list-mail-folders', { includeHiddenFolders: 'true' })).toBe('/me/mailFolders?includeHiddenFolders=true&$top=100');
    expect(await pathOf('list-mail-folders', { top: '5' })).toBe('/me/mailFolders?$top=5');
    expect(await pathOf('list-mail-child-folders', { mailFolderId: 'f1' })).toBe('/me/mailFolders/f1/childFolders?$top=100');
  });
});
