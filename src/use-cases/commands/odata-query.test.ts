import { describe, expect, it } from 'bun:test';
import { appendOData, odataQuerySchema } from './odata-query.ts';

describe('appendOData', () => {
  it('returns the path unchanged when every OData param is omitted', () => {
    expect(appendOData('/me/drives', {})).toBe('/me/drives');
  });

  it('appends $top with `?` when the path has no existing query string', () => {
    expect(appendOData('/me/drives', { top: '5' })).toBe('/me/drives?$top=5');
  });

  it('appends $top with `&` when the path already contains a query string', () => {
    expect(appendOData("/me/messages?$filter=conversationId eq 'x'", { top: '5' })).toBe("/me/messages?$filter=conversationId eq 'x'&$top=5");
  });

  it('emits params in the canonical order top, skip, select, filter, orderby, expand', () => {
    const result = appendOData('/me/messages', {
      expand: 'attachments',
      orderby: 'receivedDateTime desc',
      filter: 'isRead eq false',
      select: 'id,subject',
      skip: '10',
      top: '25',
    });
    const queryStart = result.indexOf('?');
    expect(queryStart).toBeGreaterThan(-1);
    const keys = result
      .slice(queryStart + 1)
      .split('&')
      .map((kv) => kv.split('=')[0]);
    expect(keys).toEqual(['$top', '$skip', '$select', '$filter', '$orderby', '$expand']);
  });

  it('URL-encodes the `&` and `#` characters in param values so they cannot break the query string', () => {
    const result = appendOData('/me/messages', { filter: "subject eq 'a&b#c'" });
    expect(result).toBe("/me/messages?$filter=subject%20eq%20'a%26b%23c'");
  });

  it('skips params whose value is undefined', () => {
    expect(appendOData('/me/drives', { top: '5', filter: undefined })).toBe('/me/drives?$top=5');
  });
});

describe('odataQuerySchema', () => {
  it('accepts every OData param being omitted', () => {
    const parsed = odataQuerySchema.safeParse({});
    expect(parsed.success).toBe(true);
  });

  it('accepts a positive integer string for $top', () => {
    expect(odataQuerySchema.safeParse({ top: '1' }).success).toBe(true);
    expect(odataQuerySchema.safeParse({ top: '999' }).success).toBe(true);
  });

  it('rejects $top=0 client-side because Graph itself returns badArgument on it', () => {
    expect(odataQuerySchema.safeParse({ top: '0' }).success).toBe(false);
  });

  it('rejects a non-numeric $top with a "not a number" message that names the offending value', () => {
    const parsed = odataQuerySchema.safeParse({ top: 'abc' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('not a number');
  });

  it('rejects a negative $top with a "0 and negatives" message (distinct from the non-numeric path)', () => {
    const parsed = odataQuerySchema.safeParse({ top: '-1' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('0 and negatives');
  });

  // v1.4.0 audit #7: `--top 1.5` used to say "not a number" — but 1.5 IS
  // a number, just not an integer. Detect the decimal shape distinctly so
  // the error is accurate.
  it("rejects a decimal $top (e.g. `1.5`) with a 'decimal — pagination expects whole-number counts' message, distinct from the 'not a number' path", () => {
    const parsed = odataQuerySchema.safeParse({ top: '1.5' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) {
      const msg = parsed.error.issues[0]?.message ?? '';
      expect(msg).toContain('whole-number');
      expect(msg).not.toContain('not a number');
    }
  });

  it("rejects a negative decimal $top (e.g. `-2.5`) with the same 'whole-number' message — the decimal detector wins over the signed-integer path", () => {
    const parsed = odataQuerySchema.safeParse({ top: '-2.5' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message ?? '').toContain('whole-number');
  });

  it('accepts $top=1000 (the documented Graph cap — pinning the boundary so off-by-one mutations are caught)', () => {
    expect(odataQuerySchema.safeParse({ top: '1000' }).success).toBe(true);
  });

  it('rejects $top=1001 with a "≤ 1000" message rather than letting Graph silently truncate (the audit-flagged "999999 returns 1000 with no warning" trap)', () => {
    const parsed = odataQuerySchema.safeParse({ top: '1001' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('≤ 1000');
  });

  it('rejects $top=999999 with the same cap message (the headline LLM-trap value from the audit)', () => {
    const parsed = odataQuerySchema.safeParse({ top: '999999' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('≤ 1000');
  });

  it('still accepts $skip=0 (harmless on Graph) so paging code can pass the value unchanged', () => {
    expect(odataQuerySchema.safeParse({ skip: '0' }).success).toBe(true);
  });

  it('rejects a non-numeric $skip with a "not a number" message', () => {
    const parsed = odataQuerySchema.safeParse({ skip: 'lots' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('not a number');
  });

  it('rejects a negative $skip with a "negatives" message', () => {
    const parsed = odataQuerySchema.safeParse({ skip: '-1' });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(parsed.error.issues[0]?.message).toContain('negatives');
  });

  it('rejects an empty $filter to prevent users supplying a meaningless flag', () => {
    expect(odataQuerySchema.safeParse({ filter: '' }).success).toBe(false);
  });

  it('accepts a non-empty $filter, $select, $orderby, $expand', () => {
    const parsed = odataQuerySchema.safeParse({
      filter: 'name eq foo',
      select: 'id,name',
      orderby: 'createdDateTime desc',
      expand: 'children',
    });
    expect(parsed.success).toBe(true);
  });
});
