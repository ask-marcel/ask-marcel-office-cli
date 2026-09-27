import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

type Diff = { readonly text: string; readonly added?: number; readonly removed?: number; readonly note?: string };

const text = (body: string): Result<unknown, GraphError> => ok({ contentType: 'text/plain', size: body.length, text: body });

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(graph, params);
};

const lines = (count: number, tag: string): string => Array.from({ length: count }, (_unused, i) => `row ${i} ${tag}`).join('\n');

describe('what the diff commands say when they cannot give a diff', () => {
  it('gives up at 1,000 edits, a changed line counting once each way, so 499 changed rows still diff and 501 do not', async () => {
    const pair = (count: number): ReturnType<typeof fakeGraphClient> =>
      fakeGraphClient({ get: async () => ok({ name: 'rows.md' }), getBinary: async (path: string) => text(lines(count, path.includes('/i6/') ? 'old' : 'new')) });
    const params = { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1', otherItemId: 'i7' };
    const within = await run('diff-drive-items', pair(499), params);
    if (!within.ok) throw new Error(within.error.message);
    expect(within.value).toMatchObject({ added: 499, removed: 499 });
    const beyond = await run('diff-drive-items', pair(501), params);
    if (!beyond.ok) throw new Error(beyond.error.message);
    expect((beyond.value as Diff).note).toContain('more than 1,000 added or removed lines (a changed line counts once each way)');
  });

  it('sends a version diff that gave up to the commands that can read an old version and the live file', async () => {
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'plan.md' }),
      getBinaryElevated: async () => text(lines(700, 'old')),
      getBinary: async () => text(lines(700, 'new')),
    });
    const result = await run('diff-drive-item-versions', graph, { driveId: 'd1', itemId: 'i1', versionId: '4.0' });
    if (!result.ok) throw new Error(result.error.message);
    expect((result.value as Diff).note).toContain('read the version with download-drive-item-version --format markdown and the live file with download-drive-item-as-markdown.');
  });

  it('shows a first save of an empty file as lines added, with no removed blank line', async () => {
    const graph = fakeGraphClient({ get: async () => ok({ name: 'notes.md' }), getBinary: async (path: string) => text(path.includes('/i6/') ? '' : 'Agenda\nActions') });
    const result = await run('diff-drive-items', graph, { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1', otherItemId: 'i7' });
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toMatchObject({ text: '--- a/notes.md\n+++ b/notes.md\n@@ -0,0 +1,2 @@\n+Agenda\n+Actions\n', added: 2, removed: 0 });
  });

  it('points a Loop version diff at the command that has --format original', async () => {
    const result = await run('diff-drive-item-versions', fakeGraphClient({ get: async () => ok({ name: 'standup.loop' }) }), { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('Use `download-drive-item-version --format original` for an old version');
  });

  it('says a file with no version before the instant may be newer than it, and how to read it whole', async () => {
    const graph = fakeGraphClient({
      get: async (path: string) => ok(path.includes('/versions') ? { value: [{ id: '1.0', lastModifiedDateTime: '2026-09-25T09:00:00Z' }] } : { name: 'brief.docx' }),
    });
    const result = await run('diff-drive-item-versions', graph, { driveId: 'd1', itemId: 'i1', before: '2026-09-20T00:00:00Z' });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toContain('it was created later, or its older versions were trimmed');
      expect(result.error.message).toContain('download-drive-item-as-markdown reads the file whole');
    }
  });

  it('lets --output-path land a long diff: both commands produce a body to save', () => {
    for (const name of ['diff-drive-items', 'diff-drive-item-versions']) expect(commands[name]?.meta.producesBytes).toBe(true);
  });
});
