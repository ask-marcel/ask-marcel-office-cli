import { describe, expect, it } from 'bun:test';
import { formatError } from './format-error.ts';

describe('formatError', () => {
  it('returns the message when given an Error', () => {
    expect(formatError(new Error('something broke'))).toBe('something broke');
  });

  it('returns the string as-is', () => {
    expect(formatError('raw error')).toBe('raw error');
  });

  it('converts a number to string', () => {
    expect(formatError(42)).toBe('42');
  });

  it('converts a boolean to string', () => {
    expect(formatError(false)).toBe('false');
  });

  it('stringifies a plain object', () => {
    expect(formatError({ code: 'ECONNREFUSED' })).toBe('{"code":"ECONNREFUSED"}');
  });

  it('returns fallback when JSON.stringify throws', () => {
    const circular: Record<string, unknown> = {};
    circular.self = circular;
    expect(formatError(circular)).toBe('[unstringifiable error]');
  });

  it('converts a Symbol to string via String fallback', () => {
    const sym = Symbol('test-err');
    expect(formatError(sym)).toBe('Symbol(test-err)');
  });
});
