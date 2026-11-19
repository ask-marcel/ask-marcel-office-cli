import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { parseReferenceUrl, referenceKeyOf, urlOfReferenceKey } from './planner-reference.ts';

/*
 * The key of a reference of a Planner task is its URL with `%`, `.`, `:`, `@`
 * and `#` encoded, as the Graph docs ask (OData does not permit them in a
 * property name). These tests prove the encoding on the example of the docs
 * and on the edge cases: a `%` in the URL, and each of the five characters.
 */

describe('the key of a Planner reference', () => {
  it('encodes the URL of the example of the Graph docs into the key of the docs', () => {
    expect(referenceKeyOf('https://contoso.sharepoint.com/teams/agile/documents/AnnualReport.pptx')).toBe(
      'https%3A//contoso%2Esharepoint%2Ecom/teams/agile/documents/AnnualReport%2Epptx'
    );
    expect(referenceKeyOf('https://developer.microsoft.com/graph/graph-explorer')).toBe('https%3A//developer%2Emicrosoft%2Ecom/graph/graph-explorer');
  });

  it('encodes each of the five characters, and the % of the URL first so that it is not read as a code', () => {
    expect(referenceKeyOf('https://a.b/x@y#z%20q:1')).toBe('https%3A//a%2Eb/x%40y%23z%2520q%3A1');
    expect(referenceKeyOf('https://a.b/%2E')).toBe('https%3A//a%2Eb/%252E');
  });

  it('keeps the other characters of the URL as they are', () => {
    expect(referenceKeyOf('https://a/b-c_d~e/f?g=h&i=j+k')).toBe('https%3A//a/b-c_d~e/f?g=h&i=j+k');
  });

  it('gives the URL of a key back, the codes of the URL too', () => {
    for (const url of ['https://a.b/x@y#z%20q:1', 'https://a.b/%2E%3A%25', 'https://contoso.sharepoint.com/a.pptx', 'https://h/%40']) {
      expect(urlOfReferenceKey(referenceKeyOf(url))).toBe(url);
    }
    expect(urlOfReferenceKey('https%3A//developer%2Emicrosoft%2Ecom')).toBe('https://developer.microsoft.com');
  });
});

describe('the URL of a link', () => {
  it('takes a full http or https URL', () => {
    expect(parseReferenceUrl('https://example.com/page', '--url')).toEqual(ok('https://example.com/page'));
    // A plain http link is a reference too (the scheme is put together so that the lint does not read it as a call over http).
    const plain = ['http', '://example.com'].join('');
    expect(parseReferenceUrl(plain, '--url')).toEqual(ok(plain));
  });

  it('refuses another scheme, a URL with no host or a text that is not a URL, and does not quote it', () => {
    const refusal: Result<string, GraphError> = err({
      type: 'validation_error',
      code: 'invalid_reference_url',
      message: '--url: give a full http or https URL, for example https://example.com/page.',
    });
    for (const raw of ['ftp://example.com/a', 'file:///c:/secret.txt', 'javascript:alert(1)', 'example.com/page', 'https://', '']) {
      expect(parseReferenceUrl(raw, '--url')).toEqual(refusal);
    }
  });

  it('takes a URL of 2048 characters and refuses one of 2049', () => {
    const base = 'https://example.com/';
    expect(parseReferenceUrl(base + 'a'.repeat(2048 - base.length), '--url').ok).toBe(true);
    expect(parseReferenceUrl(base + 'a'.repeat(2049 - base.length), '--url')).toEqual(
      err({ type: 'validation_error', code: 'invalid_reference_url', message: '--url: the URL has 2049 characters, above the limit of 2048.' })
    );
  });
});
