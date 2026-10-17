// The write cases of the CLI suite: a draft command run through the CLI, and the help text
// on what each write leaves behind. Moved out of cli.test.ts with their text unchanged
// (package split, step 16), so the write package takes this file whole; the helpers are
// copies of that file's.
import { describe, expect, it } from 'bun:test';
import type { AuthLadder } from '../test-helpers/auth-manager-fake.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { commands } from '../use-cases/commands/index.ts';
import { buildCli } from './cli.ts';

const captureStream = async (stream: 'stdout' | 'stderr', run: () => void | Promise<unknown>): Promise<string> => {
  const target = process[stream];
  const original = target.write.bind(target);
  let captured = '';
  const swap = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  target.write = swap;
  try {
    await run();
  } finally {
    target.write = original;
  }
  return captured;
};

const okAuth = (): AuthLadder => fakeAuthManager();

const okGraph = (value: unknown): GraphClient =>
  fakeGraphClient({
    get: async () => ({ ok: true, value }),
    post: async () => ({ ok: true, value }),
    getBinary: async () => ({ ok: true, value }),
    fetchUrl: async () => ({ ok: true, value }),
    put: async () => ({ ok: true, value }),
    delete: async () => ({ ok: true, value }),
  });

describe('buildCli command surface', () => {
  it('top-level description says what each write leaves behind, so the PDF conversions are not hidden behind a drafts-only promise', () => {
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger: createLoggerFake(), fs: createFileSystemFake() });
    const description = cli.description();
    const writeCount = Object.values(commands).filter((c) => c.meta.effect !== 'read').length;
    expect(description).toContain(`the ONLY writes are ${writeCount} commands.`);
    expect(description).toContain(
      'convert-calendar-event-attachment-to-pdf, convert-group-post-attachment-to-pdf, convert-mail-attachment-to-pdf: upload the attachment to a temporary file in the `.ask-marcel-temp` folder of your OneDrive to convert it, then try to delete that file (a failed cleanup can leave it there)'
    );
    // The cleanup is best effort, so the help must not promise that no file stays.
    expect(description).not.toContain('keeps no file');
    expect(description).not.toContain('write files');
  });

  it('sends the --comment text as the Graph reply comment on create-reply-draft', async () => {
    let captured: unknown;
    const captureGraph: GraphClient = fakeGraphClient({
      post: async (_path: string, body: unknown) => {
        captured = (body as { comment?: unknown }).comment;
        return { ok: true, value: { id: 'AAMkAGI2', isDraft: true } };
      },
    });
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger: createLoggerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', 'create-reply-draft', '--reply-to-message-id', 'AAMkAGI2', '--comment', 'Confirmed for Contoso.'])
    );
    expect(captured).toBe('Confirmed for Contoso.');
  });
});
