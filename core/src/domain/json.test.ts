import { describe, expect, it } from 'bun:test';
import { parseJson } from './json.ts';

describe('parsing JSON that may not be JSON', () => {
  it('returns the parsed value for valid JSON', () => {
    const result = parseJson('{"a":[1,2]}');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ a: [1, 2] });
  });

  it('returns an error naming the failure for invalid JSON instead of throwing', () => {
    const result = parseJson('{not json');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.length).toBeGreaterThan(0);
  });
});
