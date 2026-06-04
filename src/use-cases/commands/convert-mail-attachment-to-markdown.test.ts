import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Result } from '../../domain/result.ts';
import { buildRichOdt, buildRichPptx, buildSampleDocx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { execute } from './convert-mail-attachment-to-markdown.ts';

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

// A GraphClient whose `get`/`getBinary` are driven by URL → Result handlers.
const graphWith = (handlers: { get?: (url: string) => Result<unknown, GraphError>; getBinary?: (url: string) => Result<unknown, GraphError> }): GraphClient =>
  ({
    get: async (url: string) => handlers.get?.(url) ?? ok({}),
    post: async () => ok({}),
    getBinary: async (url: string) => handlers.getBinary?.(url) ?? ok({}),
    getElevated: async () => ok({}),
    teamsChat: async () => ok({}),
    teamsChatIc3: async () => ok({}),
    getBinaryElevated: async () => ok({}),
    fetchUrl: async () => ok({}),
    put: async () => ok({}),
    delete: async () => ok({}),
    getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  }) as GraphClient;

// A graph that returns one fileAttachment for the attachments fetch.
const fileAttachment = (name: string, bytes: Uint8Array): GraphClient =>
  graphWith({ get: () => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: toBase64(bytes) }) });

const params = { messageId: 'm1', attachmentId: 'a1' };
const asEnv = (r: Result<unknown, GraphError>): { contentType?: string; text?: string; note?: string } =>
  r.ok ? (r.value as { contentType?: string; text?: string; note?: string }) : {};

// Unconditional error assertions (no `if (error.type === ...)` guard, which would
// let a type/message/object mutant skip the assertion and survive — see LESSONS.md).
const expectApiErr = (r: Result<unknown, GraphError>, substr: string, status?: number): void => {
  expect(r.ok).toBe(false);
  if (r.ok) return;
  expect(r.error.type).toBe('api_error');
  if (status !== undefined) expect(r.error.type === 'api_error' ? r.error.status : -1).toBe(status);
  expect(r.error.message).toContain(substr);
};

describe('convert-mail-attachment-to-markdown — validation + dispatch', () => {
  it('returns a validation_error when messageId/attachmentId are missing', async () => {
    const result = await execute(graphWith({}), {});
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('propagates the graph.get error when the attachment fetch fails', async () => {
    const graph = graphWith({ get: () => ({ ok: false, error: { type: 'api_error', status: 404, message: 'no such attachment' } }) });
    expectApiErr(await execute(graph, params), 'no such attachment', 404);
  });

  it('errs when the attachment response has no @odata.type discriminator', async () => {
    expectApiErr(await execute(graphWith({ get: () => ok({ name: 'x.docx' }) }), params), 'missing @odata.type discriminator');
  });

  it('errs on an unsupported top-level attachment type', async () => {
    expectApiErr(await execute(graphWith({ get: () => ok({ '@odata.type': '#microsoft.graph.weird' }) }), params), 'unsupported attachment type: #microsoft.graph.weird');
  });
});

describe('convert-mail-attachment-to-markdown — fileAttachment formats', () => {
  it('returns the decoded text for a UTF-8 attachment, content-sniffed (no plain-text extension list)', async () => {
    const result = await execute(fileAttachment('notes.txt', new TextEncoder().encode('hello')), params);
    expect(result.ok).toBe(true);
    expect(asEnv(result).contentType).toBe('text/plain');
    expect(asEnv(result).text).toBe('hello');
  });

  it('content-sniffs a dotless-name attachment as text when its bytes are valid UTF-8 (was a 415 under the old extension list)', async () => {
    const result = await execute(fileAttachment('READMEFILE', new TextEncoder().encode('readme body')), params);
    expect(result.ok).toBe(true);
    expect(asEnv(result).text).toBe('readme body');
  });

  it('converts docx; threads includeMetadata into the DOCX metadata block when true', async () => {
    const docx = await buildSampleDocx();
    const plain = await execute(fileAttachment('r.docx', docx), params);
    expect(asEnv(plain).text).toContain('# Sample Heading');
    expect(asEnv(plain).text).not.toContain('## DOCX metadata');
    const withMeta = await execute(fileAttachment('r.docx', docx), { ...params, includeMetadata: 'true' });
    expect(asEnv(withMeta).text).toContain('## DOCX metadata');
  });

  it('converts xlsx; threads includeMetadata into the Workbook metadata block when true', async () => {
    const xlsx = buildSampleXlsx();
    expect(asEnv(await execute(fileAttachment('d.xlsx', xlsx), params)).text).toContain('## Sheet1');
    expect(asEnv(await execute(fileAttachment('d.xlsx', xlsx), { ...params, includeMetadata: 'true' })).text).toContain('## Workbook metadata');
  });

  it('rejects pptx without --include-metadata and extracts metadata with it', async () => {
    const pptx = await buildRichPptx();
    expectApiErr(await execute(fileAttachment('deck.pptx', pptx), params), 'pptx attachment not supported', 415);
    expect(asEnv(await execute(fileAttachment('deck.pptx', pptx), { ...params, includeMetadata: 'true' })).text).toContain('## PPTX metadata');
  });

  it('converts an odt body; threads includeMetadata into the OpenDocument metadata block when true', async () => {
    const odt = await buildRichOdt();
    const plain = await execute(fileAttachment('plan.odt', odt), params);
    expect(asEnv(plain).text).toContain('# Heading One');
    expect(asEnv(plain).text).not.toContain('## OpenDocument metadata');
    expect(asEnv(await execute(fileAttachment('plan.odt', odt), { ...params, includeMetadata: 'true' })).text).toContain('## OpenDocument metadata');
  });

  it('rejects a pdf attachment with the no-PDF-parser hint', async () => {
    expectApiErr(await execute(fileAttachment('scan.pdf', new Uint8Array([1])), params), 'does not bundle a PDF parser', 415);
  });

  it('rejects every image extension with the image-specific hint', async () => {
    for (const ext of ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico']) {
      expectApiErr(await execute(fileAttachment(`pic.${ext}`, new Uint8Array([1])), params), `${ext} attachment is an image and cannot be converted to markdown`, 415);
    }
  });

  it('rejects an unknown-extension BINARY attachment with the generic hint and a <no-extension> placeholder for a dotless binary name', async () => {
    expectApiErr(await execute(fileAttachment('archive.bespoke', new Uint8Array([0xff, 0xfe])), params), 'bespoke attachment not supported', 415);
    expectApiErr(await execute(fileAttachment('READMEFILE', new Uint8Array([0xff, 0xfe])), params), '<no-extension> attachment not supported', 415);
  });
});

describe('convert-mail-attachment-to-markdown — referenceAttachment', () => {
  it('errs when the referenceAttachment has no sourceUrl', async () => {
    expectApiErr(await execute(graphWith({ get: () => ok({ '@odata.type': '#microsoft.graph.referenceAttachment' }) }), params), 'referenceAttachment missing sourceUrl', 400);
  });

  it('propagates the /shares resolution error', async () => {
    const graph = graphWith({
      get: (url) =>
        url.includes('/attachments/')
          ? ok({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' })
          : { ok: false, error: { type: 'api_error', status: 403, message: 'forbidden share' } },
    });
    expectApiErr(await execute(graph, params), 'forbidden share', 403);
  });

  it('errs when the resolved driveItem is missing id or driveId', async () => {
    const graph = graphWith({
      get: (url) => (url.includes('/attachments/') ? ok({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' }) : ok({ name: 'q.docx' })),
    });
    expectApiErr(await execute(graph, params), 'resolved driveItem missing id or driveId', 500);
  });

  it('resolves a referenceAttachment through /shares and converts the linked docx', async () => {
    const docx = await buildSampleDocx();
    const graph = graphWith({
      get: (url) =>
        url.includes('/attachments/')
          ? ok({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/q.docx' })
          : ok({ id: 'i9', name: 'q.docx', parentReference: { driveId: 'd9' } }),
      getBinary: () => ok({ contentType: 'application/octet-stream', size: docx.byteLength, base64: toBase64(docx) }),
    });
    const result = await execute(graph, params);
    expect(result.ok).toBe(true);
    expect(asEnv(result).text).toContain('# Sample Heading');
  });
});

describe('convert-mail-attachment-to-markdown — itemAttachment', () => {
  const item = (inner: Record<string, unknown>): GraphClient => graphWith({ get: () => ok({ '@odata.type': '#microsoft.graph.itemAttachment', item: inner }) });

  it('errs when the itemAttachment has no inner item', async () => {
    expectApiErr(await execute(graphWith({ get: () => ok({ '@odata.type': '#microsoft.graph.itemAttachment' }) }), params), 'itemAttachment missing inner item', 400);
  });

  it('errs when the inner item has no @odata.type discriminator', async () => {
    expectApiErr(await execute(item({ subject: 'x' }), params), 'itemAttachment.item missing @odata.type discriminator', 400);
  });

  it('renders an embedded message, event, and contact to markdown', async () => {
    const message = await execute(item({ '@odata.type': '#microsoft.graph.message', subject: 'Embedded Subject', body: { content: 'hi' } }), params);
    expect(message.ok).toBe(true);
    expect(asEnv(message).text).toContain('Embedded Subject');
    const event = await execute(item({ '@odata.type': '#microsoft.graph.event', subject: 'Embedded Meeting' }), params);
    expect(asEnv(event).text).toContain('Embedded Meeting');
    const contact = await execute(item({ '@odata.type': '#microsoft.graph.contact', displayName: 'Embedded Person' }), params);
    expect(asEnv(contact).text).toContain('Embedded Person');
  });

  it('errs on an unsupported embedded item type', async () => {
    expectApiErr(await execute(item({ '@odata.type': '#microsoft.graph.task' }), params), 'unsupported embedded item type: #microsoft.graph.task', 400);
  });
});
