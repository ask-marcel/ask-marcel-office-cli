import { describe, expect, it } from 'bun:test';
import { lineDiff } from './line-diff-adapter.ts';

describe('the line diff of two texts', () => {
  it('answers a unified diff with three lines of context and the added and removed counts, without end-of-file markers', () => {
    expect(lineDiff('a/plan.md', 'one\ntwo\nthree', 'b/plan.md', 'one\n2\nthree\nfour')).toEqual({
      patch: '--- a/plan.md\n+++ b/plan.md\n@@ -1,3 +1,4 @@\n one\n-two\n+2\n three\n+four\n',
      added: 2,
      removed: 1,
    });
  });

  it('answers an empty patch for two equal texts, and nothing at all for two unrelated ones', () => {
    expect(lineDiff('a', 'same\n', 'b', 'same')).toEqual({ patch: '', added: 0, removed: 0 });
    const lines = (prefix: string): string => Array.from({ length: 600 }, (_unused, i) => `${prefix} ${i}`).join('\n');
    expect(lineDiff('a', lines('old'), 'b', lines('new'))).toBeUndefined();
  });
});
