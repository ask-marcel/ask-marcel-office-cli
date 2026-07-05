import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './convert-mail-attachment-to-pdf.ts';

// A fileAttachment whose extension routes through the temp-upload → ?format=pdf
// path (xlsx is neither plain-text, pdf, nor image), so the transform result is
// whatever `getBinary` hands back.
const xlsxAttachment = { '@odata.type': '#microsoft.graph.fileAttachment', name: 'quarter-close.xlsx', contentBytes: 'eA==' };

// get → the attachment; put → a temp item id; getBinary → the (failing) transform;
// non-attachment get → an empty temp-folder children probe for the cleanup pass.
const uploadThenTransform = (transformError: GraphError): GraphClient =>
  fakeGraphClient({
    get: async (path) => (path.includes('/attachments/') ? ok(xlsxAttachment) : ok({ value: [] })),
    put: async () => ok({ id: 'temp-item-1' }),
    getBinary: async () => err(transformError),
    delete: async () => ok({}),
  });

describe('convert-mail-attachment-to-pdf — Graph transform 406', () => {
  it('maps a 406 from ?format=pdf to an actionable transform_pdf_refused pointing at the markdown route', async () => {
    const graph = uploadThenTransform({ type: 'api_error', status: 406, message: 'InputFormatNotSupported' });

    const result = await execute(graph, { messageId: 'msg-1', attachmentId: 'att-1' });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      const error = result.error as { type: string; status?: number; code?: string; message: string };
      expect(error.type).toBe('api_error');
      expect(error.status).toBe(406);
      expect(error.code).toBe('transform_pdf_refused');
      expect(error.message).toContain('convert-mail-attachment-to-markdown');
      expect(error.message).toContain('get-mail-attachment');
    }
  });

  it('passes a non-406 transform failure through untouched — never relabels it transform_pdf_refused', async () => {
    const graph = uploadThenTransform({ type: 'api_error', status: 500, message: 'transform crashed' });

    const result = await execute(graph, { messageId: 'msg-1', attachmentId: 'att-1' });

    // Original error survives verbatim: no status rewrite, no `code` grafted on.
    expect(result).toEqual(err({ type: 'api_error', status: 500, message: 'transform crashed' }));
  });
});
