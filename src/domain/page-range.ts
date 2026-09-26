import type { Result } from './result.ts';
import { err, ok } from './result.ts';

/**
 * A selection of PDF pages such as `2`, `1-3` or `1,4,6-8`: pages count from 1,
 * a range runs low to high, parts are comma-separated and may carry spaces. It
 * lets a caller read a long scanned document a few pages at a time.
 */
type PageRange = ReadonlyArray<readonly [number, number]>;

const PART = /^(\d+)(?:-(\d+))?$/;

const parsePageRange = (raw: string): Result<PageRange, string> => {
  const spans: Array<readonly [number, number]> = [];
  for (const part of raw.split(',')) {
    const m = PART.exec(part.trim());
    const from = Number(m?.[1]);
    const to = m?.[2] === undefined ? from : Number(m[2]);
    if (m === null || from < 1 || to < from) return err(`"${raw}" is not a page selection: use page numbers from 1, ranges low to high, comma-separated, e.g. 1-3 or 1,4,6-8`);
    spans.push([from, to]);
  }
  return ok(spans);
};

const includesPage = (range: PageRange, page: number): boolean => range.some(([from, to]) => page >= from && page <= to);

export { includesPage, parsePageRange };
export type { PageRange };
