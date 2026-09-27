import { describe, expect, it } from 'bun:test';
import * as XLSX from 'xlsx';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

type Diff = { readonly text: string; readonly added?: number; readonly removed?: number; readonly note?: string };

// A weekly tracker: Budget is 3 rows × 2 cols (over a cap of 4 cells), Owners is 2 × 2 (at the cap, so rendered).
const tracker = (travel: string, owner: string): Uint8Array => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Line', 'Status'],
      ['Travel', travel],
      ['Hardware', 'open'],
    ]),
    'Budget'
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Area', 'Owner'],
      ['Finance', owner],
    ]),
    'Owners'
  );
  return new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer);
};

const asBlob = (bytes: Uint8Array): Result<unknown, GraphError> =>
  ok({ contentType: 'application/octet-stream', size: bytes.byteLength, base64: Buffer.from(bytes).toString('base64') });

const twoFiles = (week6: Uint8Array, week7: Uint8Array, name = 'tracker.xlsx'): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async () => ok({ name }),
    getBinary: async (path: string) => asBlob(path.includes('/i6/') ? week6 : week7),
  });

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<Diff> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  const result = await command.execute(graph, params);
  if (!result.ok) throw new Error(result.error.message);
  return result.value as Diff;
};

const FILES = { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1', otherItemId: 'i7' };

describe('a sheet over the --max-cells cap in a diff', () => {
  it('is never reported as unchanged: the note names the sheet that was not compared and how to compare it', async () => {
    const diff = await run('diff-drive-items', twoFiles(tracker('open', 'Robin Chen'), tracker('done', 'Robin Chen')), { ...FILES, maxCells: '4' });
    expect(diff).toMatchObject({ text: '', added: 0, removed: 0 });
    expect(diff.note).not.toContain('render to the same markdown');
    expect(diff.note).toContain('sheet Budget');
    expect(diff.note).toContain('--max-cells');
    expect(diff.note).toContain('--sheet');
    expect(diff.note).not.toContain('Owners');
  });

  it('still diffs the sheets under the cap, and says which one it left out', async () => {
    const diff = await run('diff-drive-items', twoFiles(tracker('open', 'Robin Chen'), tracker('done', 'Alex Kim')), { ...FILES, maxCells: '4' });
    expect(diff.text).toContain('-| Finance | Robin Chen |');
    expect(diff.text).toContain('+| Finance | Alex Kim |');
    expect(diff).toMatchObject({ added: 1, removed: 1 });
    expect(diff.note).toContain('sheet Budget');
  });

  it('compares one sheet alone with --sheet, cell by cell, with no left-out note', async () => {
    const diff = await run('diff-drive-items', twoFiles(tracker('open', 'Robin Chen'), tracker('done', 'Alex Kim')), { ...FILES, sheet: 'budget' });
    expect(diff.text).toContain('-| Travel | open |');
    expect(diff.text).toContain('+| Travel | done |');
    expect(diff.text).not.toContain('Owners');
    expect(diff).toMatchObject({ added: 1, removed: 1 });
    expect(diff).not.toHaveProperty('note');
  });

  it('names a capped CSV as the table, whose only remedy is a higher cap', async () => {
    const csv = (status: string): Uint8Array => new TextEncoder().encode(`line,status\nTravel,${status}\nHardware,open`);
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'tracker.csv' }),
      getBinary: async (path: string) => {
        const text = new TextDecoder().decode(csv(path.includes('/i6/') ? 'open' : 'done'));
        return ok({ contentType: 'text/plain', size: text.length, text });
      },
    });
    const diff = await run('diff-drive-items', graph, { ...FILES, maxCells: '4' });
    expect(diff.note).toContain('the table');
    expect(diff.note).not.toContain('render to the same markdown');
  });

  it('is named on the version diff too, and --sheet narrows both the version and the live file', async () => {
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'tracker.xlsx' }),
      getBinaryElevated: async () => asBlob(tracker('open', 'Robin Chen')),
      getBinary: async () => asBlob(tracker('done', 'Robin Chen')),
    });
    const capped = await run('diff-drive-item-versions', graph, { driveId: 'd1', itemId: 'i1', versionId: '4.0', maxCells: '4' });
    expect(capped.note).toContain('sheet Budget');
    expect(capped.note).not.toContain('render to the same markdown');
    const narrowed = await run('diff-drive-item-versions', graph, { driveId: 'd1', itemId: 'i1', versionId: '4.0', sheet: 'Budget' });
    expect(narrowed).toMatchObject({ added: 1, removed: 1 });
  });

  it('passes a note the conversion attached to either side, labelled with that side', async () => {
    const graph = fakeGraphClient({
      get: async () => ok({ name: 'standup.loop' }),
      getBinary: async (path: string) => {
        const html = path.includes('/i6/') ? '' : '<p>Decisions: ship Friday</p>';
        return ok({ contentType: 'text/html', size: html.length, text: html });
      },
    });
    const diff = await run('diff-drive-items', graph, FILES);
    expect(diff.text).toContain('+Decisions: ship Friday');
    expect(diff.note).toStartWith('a/standup.loop: Graph returned no HTML for this page');
  });

  it('names every table it left out, and keeps each note a sentence of its own', async () => {
    const capped = await run('diff-drive-items', twoFiles(tracker('open', 'Robin Chen'), tracker('done', 'Alex Kim')), { ...FILES, maxCells: '3' });
    expect(capped.note).toContain('Not compared: sheet Budget, sheet Owners exceeded');
    const empty = fakeGraphClient({ get: async () => ok({ name: 'standup.loop' }), getBinary: async () => ok({ contentType: 'text/html', size: 0, text: '' }) });
    const both = await run('diff-drive-items', empty, FILES);
    expect(both.note).toStartWith('The two files render to the same markdown. a/standup.loop: Graph returned no HTML');
    expect(both.note).toContain('sometimes by hours. Retry later; `list-drive-item-versions` shows when the page was last saved. b/standup.loop: Graph');
  });
});
