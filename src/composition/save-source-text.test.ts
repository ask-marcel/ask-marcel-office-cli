import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { buildCli } from './cli.ts';

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

// A reporting tool's "Excel" export: an HTML table under an .xls name. It is
// valid UTF-8, so the raw download answers it as text: the file's own bytes.
const EXPORT = '<html><body><table><tr><td>Region</td><td>Q3</td></tr><tr><td>North</td><td>12</td></tr></table></body></html>';

const exportOnDrive = fakeGraphClient({
  get: async () => ok({ name: 'report.xls' }),
  getBinary: async () => ok({ contentType: 'application/vnd.ms-excel', size: EXPORT.length, base64: Buffer.from(EXPORT).toString('base64') }),
});

const save = async (command: string, outputPath: string): Promise<{ readonly out: Record<string, unknown>; readonly saved: string | undefined }> => {
  const fs = createFileSystemFake();
  const cli = buildCli({ auth: fakeAuthManager(), graph: exportOnDrive, logger: createLoggerFake(), fs });
  const out = await captureStdout(() =>
    cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', outputPath, command, '--drive-id', 'd1', '--item-id', 'i1'])
  );
  return { out: JSON.parse(out.trim()) as Record<string, unknown>, saved: fs.snapshot(outputPath) };
};

describe('saving a text answer under a binary file name', () => {
  it("lands a raw download's text under the file's own name, since that text is the file's bytes", async () => {
    const { out, saved } = await save('download-drive-item-content', '/work/report.xls');
    expect(out).toMatchObject({ ok: true, data: { contentType: 'text/plain', savedTo: '/work/report.xls' } });
    expect(saved).toBe(EXPORT);
  });

  it('still refuses a converted render under that name, which no spreadsheet reader could open', async () => {
    const { out, saved } = await save('download-drive-item-as-markdown', '/work/report.xls');
    expect(out).toMatchObject({ ok: false, errorCode: 'text_under_binary_extension' });
    expect(saved).toBeUndefined();
  });
});
