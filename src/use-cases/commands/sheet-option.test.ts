import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { commands } from './index.ts';
import { xlsxToMarkdown } from './xlsx-to-markdown.ts';

const toBase64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');
const workbook = buildSampleXlsx();

describe('rendering one sheet of a workbook', () => {
  it('keeps only the named sheet, case-insensitively, and every sheet without the flag', async () => {
    const one = await xlsxToMarkdown(workbook, { sheet: 'sheet2' });
    if (!one.ok) throw new Error('expected ok');
    expect(one.value.text).toContain('## Sheet2');
    expect(one.value.text).not.toContain('## Sheet1');
    const all = await xlsxToMarkdown(workbook);
    if (!all.ok) throw new Error('expected ok');
    expect(all.value.text).toContain('## Sheet1');
    expect(all.value.text).toContain('## Sheet2');
  });

  it('names the sheets that exist when the requested one does not', async () => {
    const missing = await xlsxToMarkdown(workbook, { sheet: 'Totals' });
    expect(missing.ok).toBe(false);
    if (missing.ok) return;
    expect(missing.error.type).toBe('validation_error');
    expect(missing.error.message).toBe('no sheet named "Totals" in this workbook; its sheets are: Sheet1, Sheet2');
  });
});

describe('the --sheet flag on the two attachment reads', () => {
  const xlsxAttachment = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: 'plan.xlsx',
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    contentBytes: toBase64(workbook),
  };
  const textAttachment = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: 'note.txt',
    contentType: 'text/plain',
    contentBytes: toBase64(new TextEncoder().encode('hello')),
  };
  const params = { messageId: 'm1', attachmentId: 'a1' };

  for (const name of ['read-mail-attachment', 'convert-mail-attachment-to-markdown'] as const) {
    it(`${name} renders the one sheet, and refuses the flag on a file that is not a workbook`, async () => {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const one = await command.execute(fakeGraphClient({ get: async () => ok(xlsxAttachment) }), { ...params, sheet: 'Sheet1', keepQuoted: 'false', includeMetadata: 'false' });
      if (!one.ok) throw new Error('expected ok');
      const text = (one.value as { text: string }).text;
      expect(text).toContain('## Sheet1');
      expect(text).not.toContain('## Sheet2');
      const refused = await command.execute(fakeGraphClient({ get: async () => ok(textAttachment) }), { ...params, sheet: 'Sheet1' });
      expect(refused.ok).toBe(false);
      if (!refused.ok) expect(refused.error.message).toContain('--sheet applies to a workbook');
      const empty = await command.execute(fakeGraphClient({ get: async () => ok(xlsxAttachment) }), { ...params, sheet: '' });
      expect(empty.ok).toBe(false);
      expect(command.meta.options.map((o) => o.name)).toContain('sheet');
    });
  }
});

describe('--sheet on an attachment that is not a workbook file', () => {
  const params = { messageId: 'm1', attachmentId: 'a1', sheet: 'Sheet1' };
  const zipAttachment = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'reports.zip', contentType: 'application/zip', contentBytes: 'UEsFBg==' };
  const itemAttachment = { '@odata.type': '#microsoft.graph.itemAttachment', name: 'Fwd: plan', item: { '@odata.type': '#microsoft.graph.message', subject: 'plan' } };
  const loopAttachment = { '@odata.type': '#microsoft.graph.referenceAttachment', name: 'notes.loop', sourceUrl: 'https://contoso.sharepoint.com/sites/team/notes.loop' };
  const graphServing = (attachment: Record<string, unknown>): { graph: ReturnType<typeof fakeGraphClient>; paths: string[] } => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      get: async (path) => {
        paths.push(path);
        return path.startsWith('/shares/') ? ok({ id: 'i1', name: 'notes.loop', parentReference: { driveId: 'd1' } }) : ok(attachment);
      },
    });
    return { graph, paths };
  };
  const refusal = async (name: string, attachment: Record<string, unknown>): Promise<{ message: string; paths: string[] }> => {
    const command = commands[name];
    if (!command) throw new Error(`${name} is not registered`);
    const served = graphServing(attachment);
    const result = await command.execute(served.graph, params);
    if (result.ok) throw new Error(`${name} converted instead of refusing --sheet`);
    expect(result.error.type).toBe('validation_error');
    return { message: result.error.message, paths: served.paths };
  };

  it('refuses it on a zip archive instead of converting every entry', async () => {
    expect((await refusal('read-mail-attachment', zipAttachment)).message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this attachment is a zip archive');
    expect((await refusal('convert-mail-attachment-to-markdown', zipAttachment)).message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this file is a .zip');
  });

  it('refuses it on an embedded Outlook item and on a Loop page behind a link, on both reads', async () => {
    for (const name of ['read-mail-attachment', 'convert-mail-attachment-to-markdown']) {
      expect((await refusal(name, itemAttachment)).message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this attachment is an embedded Outlook item');
      const loop = await refusal(name, loopAttachment);
      expect(loop.message).toBe('--sheet applies to a workbook (xlsx, xlsm, xls); this file is a .loop');
      expect(loop.paths.some((path) => path.includes('format=html'))).toBe(false);
    }
  });
});

describe('--sheet on a legacy .xls workbook', () => {
  it('renders the one sheet of an attachment named .xls, on both reads', async () => {
    const xlsAttachment = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.xls', contentType: 'application/vnd.ms-excel', contentBytes: toBase64(workbook) };
    for (const name of ['read-mail-attachment', 'convert-mail-attachment-to-markdown'] as const) {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const one = await command.execute(fakeGraphClient({ get: async () => ok(xlsAttachment) }), { messageId: 'm1', attachmentId: 'a1', sheet: 'Sheet2' });
      if (!one.ok) throw new Error(`${name}: ${one.error.message}`);
      const text = (one.value as { text: string }).text;
      expect(text).toContain('## Sheet2');
      expect(text).not.toContain('## Sheet1');
    }
  });
});
