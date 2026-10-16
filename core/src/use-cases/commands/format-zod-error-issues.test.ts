import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { formatZodError } from './format-zod-error.ts';

/*
 * How a refused flag reads (core, package split phase 3): the flag in CLI
 * spelling, then the reason in plain words where Zod's would be jargon.
 */

const refusal = (schema: z.ZodType, value: unknown): string => {
  const parsed = schema.safeParse(value);
  return parsed.success ? 'accepted' : formatZodError(parsed.error);
};

describe('a refused flag', () => {
  it('reads "is empty" only for an empty value of a flag that needs one character', () => {
    expect(refusal(z.object({ mailFolderId: z.string().min(1) }), { mailFolderId: '' })).toBe('--mail-folder-id is empty');
    expect(refusal(z.object({ code: z.string().min(2) }), { code: 'a' })).toBe('--code Too small: expected string to have >=2 characters');
    expect(refusal(z.object({ count: z.number().min(1) }), { count: 0 })).toBe('--count is empty');
  });

  it('keeps the reason of a check that is not a minimum', () => {
    expect(refusal(z.object({ top: z.string().max(1) }), { top: 'ab' })).toBe('--top Too big: expected string to have <=1 characters');
  });

  it('names a nested flag by its path, numbers included, joined with dots', () => {
    expect(refusal(z.object({ items: z.array(z.object({ itemId: z.string() })) }), { items: [{}] })).toBe('--items.0.item-id is missing');
  });

  it('leaves out a path part that is neither a name nor an index', () => {
    const key = Symbol('key');
    const schema = z.object({ filter: z.string() }).superRefine((_value, ctx) => ctx.addIssue({ code: 'custom', message: 'refused', path: ['filter', key] }));
    expect(refusal(schema, { filter: 'x' })).toBe('--filter. refused');
  });
});
