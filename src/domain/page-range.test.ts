import { describe, expect, it } from 'bun:test';
import { includesPage, parsePageRange } from './page-range.ts';

describe('a page selection', () => {
  it('reads single pages, low-to-high ranges and comma-separated mixes, spaces allowed', () => {
    expect(parsePageRange('2')).toEqual({ ok: true, value: [[2, 2]] });
    expect(parsePageRange('1-3')).toEqual({ ok: true, value: [[1, 3]] });
    expect(parsePageRange(' 1 , 4, 6-8 ')).toEqual({
      ok: true,
      value: [
        [1, 1],
        [4, 4],
        [6, 8],
      ],
    });
  });

  it('refuses page zero, a range that runs backwards, and anything that is not page numbers, naming the rule', () => {
    for (const raw of ['0', '3-1', '1-3,', 'two', '1..3', '-2', '']) {
      const r = parsePageRange(raw);
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.error).toBe(`"${raw}" is not a page selection: use page numbers from 1, ranges low to high, comma-separated, e.g. 1-3 or 1,4,6-8`);
    }
  });

  it('holds exactly the pages it names', () => {
    const r = parsePageRange('1,4-5');
    if (!r.ok) throw new Error(r.error);
    expect([1, 2, 3, 4, 5, 6].filter((page) => includesPage(r.value, page))).toEqual([1, 4, 5]);
  });
});

describe('page numbers past nine', () => {
  it('reads multi-digit pages on both ends of a range', () => {
    expect(parsePageRange('10-12')).toEqual({ ok: true, value: [[10, 12]] });
  });
});
