import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleDocx, buildSampleXlsx, buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
import { executeLocal } from './convert-local-file-to-markdown.ts';
import { commands } from './index.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');
const textOf = (result: { ok: boolean; value?: unknown; error?: { message: string } }): string => {
  if (!result.ok) throw new Error(result.error?.message ?? 'expected ok');
  return (result.value as { text: string }).text;
};

const driveFile = (name: string, bytes: Uint8Array): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async () => ok({ name }),
    getBinary: async () => ok({ contentType: 'application/octet-stream', size: bytes.byteLength, base64: toBase64(bytes) }),
  });

describe('--sheet on download-drive-item-as-markdown', () => {
  const command = commands['download-drive-item-as-markdown'];
  const params = { driveId: 'd1', itemId: 'i1' };

  it('renders the one sheet of a workbook, and names the sheets when the name is unknown', async () => {
    if (!command) throw new Error('download-drive-item-as-markdown is not registered');
    const one = textOf(await command.execute(driveFile('plan.xlsx', buildSampleXlsx()), { ...params, sheet: 'sheet2' }));
    expect(one).toContain('## Sheet2');
    expect(one).not.toContain('## Sheet1');
    const missing = await command.execute(driveFile('plan.xlsx', buildSampleXlsx()), { ...params, sheet: 'Totals' });
    expect(missing.ok).toBe(false);
    if (!missing.ok) expect(missing.error.message).toBe('no sheet named "Totals" in this workbook; its sheets are: Sheet1, Sheet2');
    expect(command.meta.options.map((o) => o.name)).toContain('sheet');
  });

  it('refuses the flag on a document and on a Loop page', async () => {
    if (!command) throw new Error('download-drive-item-as-markdown is not registered');
    const docx = await command.execute(driveFile('brief.docx', await buildSampleDocx()), { ...params, sheet: 'Sheet1' });
    expect(docx.ok).toBe(false);
    if (!docx.ok) expect(docx.error.message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this file is a .docx');
    const loop = await command.execute(driveFile('notes.loop', new Uint8Array()), { ...params, sheet: 'Sheet1' });
    expect(loop.ok).toBe(false);
    if (!loop.ok) expect(loop.error.message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this file is a .loop');
  });
});

describe('--sheet on convert-local-file-to-markdown', () => {
  const local = async (path: string, bytes: Uint8Array, extra: Record<string, string>): Promise<ReturnType<typeof executeLocal>> => {
    const fs = createFileSystemFake();
    fs.seedBytes(path, bytes);
    return executeLocal(fs, { path, ...extra });
  };

  it('renders the one sheet of a local workbook', async () => {
    const one = textOf(await local('/work/call-log.xlsx', buildSampleXlsx(), { sheet: 'Sheet1' }));
    expect(one).toContain('## Sheet1');
    expect(one).not.toContain('## Sheet2');
    expect(commands['convert-local-file-to-markdown']?.meta.options.map((o) => o.name)).toContain('sheet');
  });

  it('refuses the flag on a zip archive, and applies --max-cells to the workbooks inside one', async () => {
    const zip = await buildSampleZipArchive();
    const refused = await local('/work/reports.zip', zip, { sheet: 'Sheet1' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this file is a zip archive');
    const capped = await local('/work/reports.zip', zip, { maxCells: '4' });
    if (!capped.ok) throw new Error(capped.error.message);
    expect(JSON.stringify(capped.value)).not.toContain('Alice');
  });
});
