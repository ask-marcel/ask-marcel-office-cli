import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

// A 3 × 2 tracker: over a cap of 4 cells, under a cap of 6.
const CSV = 'line,status\nTravel,open\nHardware,done';
const b64 = (bytes: Uint8Array | string): string => Buffer.from(bytes).toString('base64');
const zipOf = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('tracker.csv', CSV);
  return zip.generateAsync({ type: 'uint8array' });
};

const run = async (name: string, graph: ReturnType<typeof fakeGraphClient>, params: Record<string, string>): Promise<string> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  const result = await command.execute(graph, params);
  if (!result.ok) throw new Error(result.error.message);
  return JSON.stringify(result.value);
};

const fileAttachment = (name: string, bytes: Uint8Array | string): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({ get: async () => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: b64(bytes) }) });

type Case = { readonly name: string; readonly graph: () => Promise<ReturnType<typeof fakeGraphClient>>; readonly params: Record<string, string> };

const CASES: ReadonlyArray<Case> = [
  { name: 'convert-mail-attachment-zip-to-markdown', graph: async () => fileAttachment('pack.zip', await zipOf()), params: { messageId: 'm1', attachmentId: 'a1' } },
  {
    name: 'convert-drive-item-zip-to-markdown',
    graph: async () => {
      const zip = await zipOf();
      return fakeGraphClient({ getBinary: async () => ok({ contentType: 'application/zip', size: zip.byteLength, base64: b64(zip) }) });
    },
    params: { driveId: 'd1', itemId: 'i1' },
  },
  { name: 'convert-calendar-event-attachment-to-markdown', graph: async () => fileAttachment('tracker.csv', CSV), params: { eventId: 'e1', attachmentId: 'a1' } },
  {
    name: 'convert-group-post-attachment-to-markdown',
    graph: async () => fileAttachment('tracker.csv', CSV),
    params: { groupId: 'g1', threadId: 't1', postId: 'p1', attachmentId: 'a1' },
  },
  {
    name: 'download-drive-item-version',
    graph: async () =>
      fakeGraphClient({ get: async () => ok({ name: 'tracker.csv' }), getBinaryElevated: async () => ok({ contentType: 'text/plain', size: CSV.length, text: CSV }) }),
    params: { driveId: 'd1', itemId: 'i1', versionId: '2.0', format: 'markdown' },
  },
];

describe('--max-cells on every command whose cap hint names it', () => {
  for (const c of CASES) {
    it(`${c.name} caps a sheet at --max-cells and renders it under a higher one`, async () => {
      expect(commands[c.name]?.meta.options.map((o) => o.name)).toContain('max-cells');
      expect(await run(c.name, await c.graph(), { ...c.params, maxCells: '4' })).toContain('Table omitted');
      const rendered = await run(c.name, await c.graph(), { ...c.params, maxCells: '6' });
      expect(rendered).toContain('| Travel | open |');
      expect(rendered).not.toContain('Table omitted');
    });
  }

  it('refuses a cap that is not a positive whole number', async () => {
    const command = commands['convert-calendar-event-attachment-to-markdown'];
    if (!command) throw new Error('not registered');
    const refused = await command.execute(fileAttachment('tracker.csv', CSV), { eventId: 'e1', attachmentId: 'a1', maxCells: '0' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('validation_error');
  });
});
