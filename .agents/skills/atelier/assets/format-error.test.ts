/*
 * Test for the shipped formatError helper. Copy it alongside format-error.ts
 * into src/domain/utilities/ so the file survives the mutation gate (it lives in
 * src/domain/**, which Stryker mutates and the 90% break threshold enforces).
 *
 * The non-finite-number case is load-bearing, not decorative: it is the only
 * input that distinguishes the `typeof err === 'number'` branch from the
 * JSON.stringify fallback (String(NaN) is 'NaN'; JSON.stringify(NaN) is 'null').
 * Without it, every mutant on that branch is equivalent and survives.
 */
import { describe, expect, test } from 'bun:test';
import { formatError } from './format-error.ts';

describe('formatError renders any thrown value as a readable string', () => {
  test('an Error yields its message', () => {
    expect(formatError(new Error('boom'))).toBe('boom');
  });

  test('a raw string is returned unchanged', () => {
    expect(formatError('plain failure')).toBe('plain failure');
  });

  test('a finite number is stringified', () => {
    expect(formatError(42)).toBe('42');
  });

  test('non-finite numbers keep their String() form, not JSON null', () => {
    expect(formatError(Number.NaN)).toBe('NaN');
    expect(formatError(Number.POSITIVE_INFINITY)).toBe('Infinity');
  });

  test('a plain object is JSON-encoded', () => {
    expect(formatError({ code: 'E_IO', retriable: true })).toBe('{"code":"E_IO","retriable":true}');
  });

  test('a value that cannot be stringified falls back to a placeholder', () => {
    const circular: { self?: unknown } = {};
    circular.self = circular;
    expect(formatError(circular)).toBe('[unstringifiable error]');
  });
});
