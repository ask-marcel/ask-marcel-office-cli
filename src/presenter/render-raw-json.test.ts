import { describe, expect, it } from 'bun:test';
import { renderErrorToString, renderToString } from './render-to-string.ts';

describe('the raw-json output', () => {
  it('prints the payload alone, without the envelope or the paging cursors', () => {
    const body = { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=10', '@odata.deltaLink': 'd', '@odata.count': 1 };
    expect(renderToString(body, 'raw-json')).toBe('{"value":[{"id":"a"}]}\n');
    expect(renderToString([1, 2], 'raw-json')).toBe('[1,2]\n');
    expect(renderToString(undefined, 'raw-json')).toBe('null\n');
  });

  it('adds no size hint, however large the payload', () => {
    const big = { text: 'x'.repeat(60_000) };
    expect(
      renderToString(big, 'raw-json', { commandName: 'download-drive-item-as-markdown', producesBytes: true, supportsSelect: false, supportsTop: false, surface: 'cli' })
    ).toBe(`${JSON.stringify(big)}\n`);
  });

  it('still prints a failure as the JSON error envelope, so a script can tell it from data', () => {
    expect(renderErrorToString('itemNotFound: gone', 'raw-json', 'itemNotFound', 'graph', 5)).toBe(renderErrorToString('itemNotFound: gone', 'json', 'itemNotFound', 'graph', 5));
    expect(JSON.parse(renderErrorToString('boom', 'raw-json'))).toMatchObject({ ok: false, error: 'boom' });
  });
});
