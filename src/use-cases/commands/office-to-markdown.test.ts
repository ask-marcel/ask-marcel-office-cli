import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import {
  buildLegacyXls,
  buildPdfNoImages,
  buildPdfWithText,
  buildRichOdt,
  buildRichPptx,
  buildSampleDoc,
  buildSampleDocx,
  buildSampleXlsx,
} from '../../test-helpers/office-fixtures.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

const noopGraph = (overrides: Partial<GraphClient>): GraphClient => ({
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
  ...overrides,
});

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

// 0xff/0xfe never start a valid UTF-8 sequence; 0x80/0x81 are stray continuation
// bytes — so these content-sniff as binary (used where a non-Office file must hit
// the hint rather than the text fallback).
const BINARY_BYTES = new Uint8Array([0xff, 0xfe, 0xfd, 0x80, 0x81]);
const bytesGraph = (bytes: Uint8Array, contentType = 'application/octet-stream'): GraphClient =>
  noopGraph({ getBinary: async () => ok({ contentType, size: bytes.byteLength, base64: toBase64(bytes) }) });

describe('officeToMarkdown — extension dispatch', () => {
  it('inlines plain-text source bytes (txt/md/json/etc.) as { contentType: "text/plain", size, text } when Graph returns them directly', async () => {
    const graph = noopGraph({
      getBinary: async (path) => {
        expect(path).toBe('/drives/d1/items/i1/content');
        return ok({ contentType: 'text/plain', size: 5, text: 'hello' });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'notes.txt');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ contentType: 'text/plain', size: 5, text: 'hello' });
    }
  });

  it('inlines plain-text source bytes by following a CDN downloadUrl redirect (the common large-file case the audit flagged as still URL-only)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/notes.txt' }),
      fetchUrl: async (url) => {
        expect(url).toBe('https://contoso.sharepoint.com/cdn/notes.txt');
        return ok({ contentType: 'text/plain', size: 13, text: 'hello via CDN' });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'notes.txt');
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toEqual({ contentType: 'text/plain', size: 13, text: 'hello via CDN' });
    }
  });

  it('routes docx through mammoth → turndown via the docx-to-markdown helper', async () => {
    const docxBytes = await buildSampleDocx();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: docxBytes.byteLength, base64: toBase64(docxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('# Sample Heading');
    }
  });

  it('routes xlsx through sheetjs → markdown table per sheet via the xlsx-to-markdown helper', async () => {
    const xlsxBytes = buildSampleXlsx();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: xlsxBytes.byteLength, base64: toBase64(xlsxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('## Sheet1');
      expect(env.text).toContain('## Sheet2');
    }
  });

  it('routes loop/fluid/wbtx/whiteboard through the existing Graph ?format=html pipeline (the four inputs Microsoft documents)', async () => {
    let calledPath = '';
    const graph = noopGraph({
      getBinary: async (path) => {
        calledPath = path;
        return ok({ contentType: 'text/html', size: 13, text: '<p>loop x</p>' });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'planning.loop');
    expect(result.ok).toBe(true);
    expect(calledPath).toBe('/drives/d1/items/i1/content?format=html');
  });

  it('extracts pptx slide text (titles, body, speaker notes inline) by default — no --include-metadata required', async () => {
    const graph = bytesGraph(await buildRichPptx());
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'deck.pptx');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; text: string };
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toContain('## Slide 1');
    expect(env.text).toContain('Quarterly Review'); // title
    expect(env.text).toContain('see the portal'); // body shape text
    expect(env.text).toContain('Remember to mention the Q3 shortfall'); // speaker notes inline
    expect(env.text).not.toContain('## PPTX metadata'); // side-channel only with the flag
  });

  it('aliases the macro-enabled / template families onto their base parser (.docm → docx, .xlsm → xlsx, .pptm → pptx metadata)', async () => {
    const docxBytes = await buildSampleDocx();
    const docmGraph = noopGraph({ getBinary: async () => ok({ contentType: 'application/octet-stream', size: docxBytes.byteLength, base64: toBase64(docxBytes) }) });
    const docm = await officeToMarkdown(docmGraph, '/drives/d1/items/i1/content', 'report.docm');
    expect(docm.ok).toBe(true);
    if (docm.ok) expect((docm.value as { text: string }).text).toContain('# Sample Heading');

    const xlsxBytes = buildSampleXlsx();
    const xlsmGraph = noopGraph({ getBinary: async () => ok({ contentType: 'application/octet-stream', size: xlsxBytes.byteLength, base64: toBase64(xlsxBytes) }) });
    const xlsm = await officeToMarkdown(xlsmGraph, '/drives/d1/items/i1/content', 'data.xlsm');
    expect(xlsm.ok).toBe(true);
    if (xlsm.ok) expect((xlsm.value as { text: string }).text).toContain('## Sheet1');

    const pptxBytes = await buildRichPptx();
    const pptmGraph = noopGraph({ getBinary: async () => ok({ contentType: 'application/octet-stream', size: pptxBytes.byteLength, base64: toBase64(pptxBytes) }) });
    const pptm = await officeToMarkdown(pptmGraph, '/drives/d1/items/i1/content', 'deck.pptm', { includeMetadata: true });
    expect(pptm.ok).toBe(true);
    if (pptm.ok) expect((pptm.value as { text: string }).text).toContain('## PPTX metadata');
  });

  it('appends the PPTX metadata block after the slide text when --include-metadata true is set', async () => {
    const pptxBytes = await buildRichPptx();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: pptxBytes.byteLength, base64: toBase64(pptxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'deck.pptx', { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('## Slide 1'); // slide body present
      expect(env.text).toContain('Quarterly Review');
      expect(env.text).toContain('## PPTX metadata'); // + side-channel block
    }
  });

  it('routes odt through the OpenDocument body converter and appends the metadata block when --include-metadata true is set', async () => {
    const odtBytes = await buildRichOdt();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: odtBytes.byteLength, base64: toBase64(odtBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'plan.odt', { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('# Heading One');
      expect(env.text).toContain('## OpenDocument metadata');
      expect(env.text).toContain('Q4 Plan');
    }
  });

  it('converts the odt body to markdown (no metadata block) even when --include-metadata is not set', async () => {
    const odtBytes = await buildRichOdt();
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'application/octet-stream', size: odtBytes.byteLength, base64: toBase64(odtBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'plan.odt');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('# Heading One');
      expect(env.text).not.toContain('## OpenDocument metadata');
    }
  });

  it('errs with the generic 38-input-set hint for a non-pdf extension Graph accepts on the pdf path (rtf, etc.)', async () => {
    const graph = bytesGraph(BINARY_BYTES);
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'notes.rtf');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('rtf not supported');
    expect(result.error.message).toContain('38 input extensions');
  });

  it('extracts a born-digital PDF’s text layer as a text/plain envelope', async () => {
    const graph = bytesGraph(buildPdfWithText(), 'application/pdf');
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.pdf');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType: string; text: string };
    expect(v.contentType).toBe('text/plain');
    expect(v.text).toContain('Hello from the');
  });

  it('errs 415 with a vision-model hint for a scanned / image-only PDF (no text layer)', async () => {
    const graph = bytesGraph(buildPdfNoImages(), 'application/pdf');
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'scan.pdf');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('no extractable text layer');
    expect(result.error.message).toContain('download-drive-item-as-pdf');
  });

  it('propagates a getBinary api_error from the pdf path', async () => {
    const graph = noopGraph({ getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 404, message: 'pdf not found' } }) });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.pdf');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('pdf not found');
  });

  it('converts a legacy .xls (BIFF / OLE) through the sheetjs pipeline to a markdown table', async () => {
    const graph = bytesGraph(buildLegacyXls());
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'budget.xls');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; text: string };
    expect(env.contentType).toBe('text/markdown');
    expect(env.text).toContain('## Legacy'); // the sheet name
    expect(env.text).toContain('Alice');
  });

  it('extracts a legacy .doc (OLE binary) body as a text/plain envelope', async () => {
    const graph = bytesGraph(await buildSampleDoc());
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'memo.doc');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; text: string };
    expect(env.contentType).toBe('text/plain');
    expect(env.text).toContain('Hello from the legacy doc');
  });

  it('propagates a getBinary api_error from the legacy .xls and .doc paths', async () => {
    const xlsGraph = noopGraph({ getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 403, message: 'xls forbidden' } }) });
    const xls = await officeToMarkdown(xlsGraph, '/drives/d1/items/i1/content', 'old.xls');
    expect(xls.ok).toBe(false);
    expect(!xls.ok && xls.error.type === 'api_error' ? xls.error.message : '').toContain('xls forbidden');
    const docGraph = noopGraph({ getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 403, message: 'doc forbidden' } }) });
    const doc = await officeToMarkdown(docGraph, '/drives/d1/items/i1/content', 'old.doc');
    expect(doc.ok).toBe(false);
    expect(!doc.ok && doc.error.type === 'api_error' ? doc.error.message : '').toContain('doc forbidden');
  });

  it('errs 415 for a legacy .ppt pointing at the convert-to-PDF-first workflow', async () => {
    const graph = bytesGraph(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])); // OLE bytes are fetched, then the ppt branch errs
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'deck.ppt');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('ppt (legacy PowerPoint');
    expect(result.error.message).toContain('download-drive-item-as-pdf');
  });

  it('a known-binary extension (zip, mp4, …) is fetched and content-sniffed, then errs 415 with the generic *-as-pdf hint — no dedicated short-circuit list', async () => {
    const graph = bytesGraph(BINARY_BYTES);
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'sources.zip');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('zip not supported');
    expect(result.error.message).toContain('38 input extensions');
  });

  it('errs 415 with the image hint for a raster image (png), pointing at extract-drive-item-images / a vision model', async () => {
    const graph = bytesGraph(new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'photo.png');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('png is an image');
    expect(result.error.message).toContain('extract-drive-item-images');
  });

  it('an .svg (which is XML text) content-sniffs to text/plain rather than being rejected as binary', async () => {
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><text>org chart label</text></svg>');
    const graph = bytesGraph(svg);
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'diagram.svg');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; text: string };
    expect(env.contentType).toBe('text/plain');
    expect(env.text).toContain('<text>org chart label</text>');
  });

  it('threads includeMetadata into the docx and xlsx converters (appends the metadata block when true)', async () => {
    const docxBytes = await buildSampleDocx();
    const docxGraph = noopGraph({ getBinary: async () => ok({ contentType: 'application/octet-stream', size: docxBytes.byteLength, base64: toBase64(docxBytes) }) });
    const docx = await officeToMarkdown(docxGraph, '/drives/d1/items/i1/content', 'report.docx', { includeMetadata: true });
    expect(docx.ok).toBe(true);
    if (docx.ok) expect((docx.value as { text: string }).text).toContain('## DOCX metadata');

    const xlsxBytes = buildSampleXlsx();
    const xlsxGraph = noopGraph({ getBinary: async () => ok({ contentType: 'application/octet-stream', size: xlsxBytes.byteLength, base64: toBase64(xlsxBytes) }) });
    const xlsx = await officeToMarkdown(xlsxGraph, '/drives/d1/items/i1/content', 'data.xlsx', { includeMetadata: true });
    expect(xlsx.ok).toBe(true);
    if (xlsx.ok) expect((xlsx.value as { text: string }).text).toContain('## Workbook metadata');
  });

  it('errs with `<no-extension>` placeholder when the filename has no dot', async () => {
    const graph = bytesGraph(BINARY_BYTES);
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'README');
    // Unconditional assertions (no `type === 'api_error'` guard, which would let an
    // ObjectLiteral / StringLiteral mutant on the error envelope skip the check — see LESSONS.md).
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('<no-extension>');
  });

  it('errs with `<no-extension>` placeholder when the filename ends with a dot but no extension', async () => {
    const graph = bytesGraph(BINARY_BYTES);
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'oddfile.');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toContain('<no-extension>');
  });

  it('returns any UTF-8-decodable file as text/plain via content-sniffing — any extension, none, or an unlisted one (no plain-text list)', async () => {
    const textBytes = new TextEncoder().encode('server {\n  listen 80;\n}\n');
    const graph = bytesGraph(textBytes);
    for (const name of ['nginx.conf', 'README', 'notes.txt', 'data.weirdext']) {
      const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', name);
      expect(result.ok).toBe(true);
      if (!result.ok) continue;
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toContain('listen 80');
    }
  });

  it('propagates getBinary api_error from the docx path', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 404, message: 'not found' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('propagates getBinary api_error from the xlsx path', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 401, message: 'unauthorized' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(401);
    }
  });

  it('errs cleanly when neither getBinary nor fetchUrl produces base64 / text bytes (defensive guard for unknown shapes)', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ surprise: true }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unexpected envelope');
    }
  });

  it('follows a `@microsoft.graph.downloadUrl` 302 redirect to fetch docx bytes (the real shape Graph returns for /content)', async () => {
    const docxBytes = await buildSampleDocx();
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/abc.docx' }),
      fetchUrl: async (url) => {
        expect(url).toBe('https://contoso.sharepoint.com/cdn/abc.docx');
        return ok({ contentType: 'application/octet-stream', size: docxBytes.byteLength, base64: toBase64(docxBytes) });
      },
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('# Sample Heading');
    }
  });

  it('follows a `@microsoft.graph.downloadUrl` 302 redirect to fetch xlsx bytes', async () => {
    const xlsxBytes = buildSampleXlsx();
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/data.xlsx' }),
      fetchUrl: async () => ok({ contentType: 'application/octet-stream', size: xlsxBytes.byteLength, base64: toBase64(xlsxBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { text: string };
      expect(env.text).toContain('## Sheet1');
    }
  });

  it('propagates an err from fetchUrl when following the downloadUrl redirect fails', async () => {
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/x' }),
      fetchUrl: async () => ({ ok: false, error: { type: 'network_error' as const, message: 'socket reset' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'report.docx');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'network_error') {
      expect(result.error.message).toBe('socket reset');
    }
  });

  it('routes csv through csvToMarkdownTable (real markdown table) — follows the 302 redirect and decodes as text', async () => {
    const csv = 'Name,Score\nAlice,42\nBob,7';
    const graph = noopGraph({
      getBinary: async () => ok({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/data.csv' }),
      fetchUrl: async () => ok({ contentType: 'text/csv', size: csv.length, text: csv }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.csv');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { contentType: string; text: string };
      expect(env.contentType).toBe('text/markdown');
      expect(env.text).toContain('| Name | Score |');
      expect(env.text).toContain('| Alice | 42 |');
    }
  });

  it('routes csv through csvToMarkdownTable when Graph returns the bytes inline (no 302) by decoding base64', async () => {
    const csv = 'a,b\n1,2';
    const csvBytes = new TextEncoder().encode(csv);
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/csv', size: csvBytes.byteLength, base64: toBase64(csvBytes) }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.csv');
    expect(result.ok).toBe(true);
    if (result.ok) {
      const env = result.value as { text: string };
      expect(env.text).toContain('| a | b |');
      expect(env.text).toContain('| 1 | 2 |');
    }
  });

  it('applies the --max-cells cap on the standalone .csv path (a large csv no longer OOMs — audit A2)', async () => {
    const csv = 'a,b,c\nd,e,f\ng,h,i';
    const csvBytes = new TextEncoder().encode(csv);
    const graph = noopGraph({ getBinary: async () => ok({ contentType: 'text/csv', size: csvBytes.byteLength, base64: toBase64(csvBytes) }) });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.csv', { maxCells: 4 });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { text: string };
    expect(env.text).not.toContain('| --- |'); // table omitted
    expect(env.text).toContain('get-excel-range'); // band-by-band hint instead
  });

  it('propagates getBinary api_error from the csv path', async () => {
    const graph = noopGraph({
      getBinary: async () => ({ ok: false, error: { type: 'api_error' as const, status: 403, message: 'forbidden' } }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.csv');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(403);
    }
  });

  it('reads the text envelope path of getBinary for xlsx (rare but valid for very small workbooks)', async () => {
    // synthetic: xlsx text envelope cannot really happen, but the decoder must still cover the text branch.
    const graph = noopGraph({
      getBinary: async () => ok({ contentType: 'text/plain', size: 4, text: 'PK' }),
    });
    const result = await officeToMarkdown(graph, '/drives/d1/items/i1/content', 'data.xlsx');
    // sheetjs will reject; we only need to verify the text-decoder branch was exercised.
    expect(result.ok).toBe(false);
  });
});
