import { describe, expect, it } from 'bun:test';
import { andThen, err, isErr, isOk, map, mapErr, ok, unwrap, unwrapOr } from './result.ts';

describe('result utilities', () => {
  it('unwrap returns value on Ok', () => {
    expect(unwrap(ok(42))).toBe(42);
  });

  it('unwrap throws on Err', () => {
    expect(() => unwrap(err('oops'))).toThrow('oops');
  });

  it('unwrapOr returns value on Ok, fallback on Err', () => {
    expect(unwrapOr(ok(42), 0)).toBe(42);
    expect(unwrapOr(err('no'), 0)).toBe(0);
  });

  it('map transforms value on Ok, passes Err through', () => {
    expect(map(ok(1), (n) => n * 2)).toEqual(ok(2));
    expect(map(err('no') as ReturnType<typeof ok<number>> | ReturnType<typeof err<string>>, (n: number) => n * 2)).toEqual(err('no'));
  });

  it('mapErr transforms error on Err, passes Ok through', () => {
    expect(mapErr(err('a'), (e) => `wrapped: ${e}`)).toEqual(err('wrapped: a'));
    expect(mapErr(ok(1), (e) => e)).toEqual(ok(1));
  });

  it('andThen chains results on Ok, short-circuits on Err', () => {
    expect(andThen(ok(1), (n) => ok(n + 1))).toEqual(ok(2));
    expect(andThen(err('no') as ReturnType<typeof ok<number>> | ReturnType<typeof err<string>>, () => ok(2))).toEqual(err('no'));
  });

  it('isOk and isErr narrow correctly', () => {
    const good = ok('yes');
    const bad = err('no');
    expect(isOk(good)).toBe(true);
    expect(isErr(good)).toBe(false);
    expect(isOk(bad)).toBe(false);
    expect(isErr(bad)).toBe(true);
  });
});
