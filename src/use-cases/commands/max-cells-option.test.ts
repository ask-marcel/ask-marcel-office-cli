import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleXlsx, buildSampleZipArchive } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';
import { csvToMarkdownSection } from './xlsx-to-markdown.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');
const params = { messageId: 'm1', attachmentId: 'a1' };
const xlsxAttachment = {
  '@odata.type': '#microsoft.graph.fileAttachment',
  name: 'call-log.xlsx',
  contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  contentBytes: toBase64(buildSampleXlsx()),
};

const textOf = (result: { ok: boolean; value?: unknown }): string => {
  if (!result.ok) throw new Error('expected ok');
  return JSON.stringify(result.value);
};

describe('--max-cells on the two mail-attachment reads', () => {
  for (const name of ['read-mail-attachment', 'convert-mail-attachment-to-markdown'] as const) {
    it(`${name} caps a sheet over the limit to a hint, renders it whole under a raised limit, and refuses a limit that is not a positive integer`, async () => {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const graph = fakeGraphClient({ get: async () => ok(xlsxAttachment) });
      const capped = textOf(await command.execute(graph, { ...params, maxCells: '4' }));
      expect(capped).toContain('## Sheet1');
      expect(capped).not.toContain('Alice');
      expect(capped).toContain('Raise the cap with `--max-cells <N>`');
      expect(textOf(await command.execute(graph, { ...params, maxCells: '100' }))).toContain('Alice');
      for (const maxCells of ['0', 'x5', '5x']) {
        const refused = await command.execute(graph, { ...params, maxCells });
        expect(refused.ok).toBe(false);
        if (!refused.ok) expect(refused.error.message).toContain('must be a positive integer');
      }
      expect(command.meta.options.map((o) => o.name)).toContain('max-cells');
    });
  }

  it('read-mail-attachment applies the limit to the workbooks inside a zip attachment', async () => {
    const zipAttachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'reports.zip',
      contentType: 'application/zip',
      contentBytes: toBase64(await buildSampleZipArchive()),
    };
    const command = commands['read-mail-attachment'];
    if (!command) throw new Error('read-mail-attachment is not registered');
    const graph = fakeGraphClient({ get: async () => ok(zipAttachment) });
    expect(textOf(await command.execute(graph, { ...params, maxCells: '4' }))).not.toContain('Alice');
    expect(textOf(await command.execute(graph, params))).toContain('Alice');
  });
});

describe('the hint that replaces a sheet over the cap', () => {
  it('names what every source can do, then the band-by-band route a workbook in a drive also has', () => {
    expect(csvToMarkdownSection('Ocean', 'a,b\n1,2\n3,4', 2)).toBe(
      "## Ocean\n\n> _Table omitted: this sheet's used range is ~6 cells (3 rows × 2 cols), over the `--max-cells` 2 render cap. Raise the cap with `--max-cells <N>`, with `--output-path` to land a large render on disk; a workbook in OneDrive or SharePoint also reads band-by-band through `get-excel-used-range`, then `get-excel-range --address 'A1:Cn'` per band._"
    );
  });
});
