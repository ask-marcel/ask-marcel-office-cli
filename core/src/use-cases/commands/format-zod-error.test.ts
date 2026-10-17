import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { formatZodError } from './format-zod-error.ts';

describe('formatZodError', () => {
  it('translates a missing-field error to `--<flag> is missing` keyed off the kebab-cased schema path', () => {
    const result = z.object({ messageRuleId: z.string().min(1) }).safeParse({});
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(formatZodError(result.error)).toBe('--message-rule-id is missing');
    }
  });

  it('translates a min(1) empty-string error to `--<flag> is empty`', () => {
    const result = z.object({ query: z.string().min(1) }).safeParse({ query: '' });
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(formatZodError(result.error)).toBe('--query is empty');
    }
  });

  it('joins multiple flag failures with `; ` so multi-field validation stays on one line', () => {
    const result = z.object({ siteId: z.string().min(1), listId: z.string().min(1) }).safeParse({ siteId: '', listId: '' });
    expect(result.success).toBe(false);
    if (!result.success) {
      const formatted = formatZodError(result.error);
      expect(formatted).toContain('--site-id is empty');
      expect(formatted).toContain('--list-id is empty');
      expect(formatted).toContain('; ');
    }
  });

  it("falls back to Zod's own message when the issue is neither invalid_type nor too_small min=1", () => {
    const result = z.object({ top: z.string().regex(/^\d+$/, 'must be a non-negative integer') }).safeParse({ top: 'lots' });
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(formatZodError(result.error)).toBe('--top must be a non-negative integer');
    }
  });

  it('uses `<root>` as the path label for issues without a path (root-level refinements)', () => {
    const schema = z.string().refine(() => false, { message: 'always fails' });
    const result = schema.safeParse('anything');
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(formatZodError(result.error)).toBe('<root>: always fails');
    }
  });

  it('returns "validation failed" when the ZodError has zero issues', () => {
    const fakeError = { issues: [] } as unknown as z.ZodError;
    expect(formatZodError(fakeError)).toBe('validation failed');
  });
});
