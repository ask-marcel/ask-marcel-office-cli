import { describe, expect, it } from 'bun:test';
import {
  filterSelectOptions,
  noSkipOptions,
  noSkipShape,
  odataQueryOptions,
  odataQuerySchema,
  pickODataOptions,
  pickODataShape,
  selectExpandOptions,
  selectOnlyOptions,
  selectOnlyShape,
  topOnlyOptions,
  topOnlyShape,
} from './odata-query.ts';

/*
 * The OData passthroughs every list builder advertises (core, package split
 * phase 3): the flags each subset offers, and the message a bad `--top` or
 * `--skip` gets.
 */

const messageFor = (params: Record<string, string>): string | undefined => {
  const parsed = odataQuerySchema.safeParse(params);
  return parsed.success ? undefined : parsed.error.issues.map((issue) => issue.message).join('; ');
};

const NOT_A_NUMBER = (value: string): string => `must be a positive integer (got "${value}", which is not a number)`;
const DECIMAL = (value: string): string => `must be a positive integer (got "${value}", which is a decimal — pagination expects whole-number counts)`;

describe('a bad --top', () => {
  it('is named a decimal only when it is a whole decimal number', () => {
    expect(['1.5', '-2.5', '12.5', '1.25'].map((value) => messageFor({ top: value }))).toEqual(['1.5', '-2.5', '12.5', '1.25'].map(DECIMAL));
    expect(['x1.5', '1.5x'].map((value) => messageFor({ top: value }))).toEqual(['x1.5', '1.5x'].map(NOT_A_NUMBER));
  });

  it('is named zero or negative only when it is a whole integer', () => {
    expect(['0', '-5', '-12'].map((value) => messageFor({ top: value }))).toEqual(
      Array.from({ length: 3 }, () => 'must be a positive integer (Graph rejects top=0 and negatives)')
    );
    expect(['x5', '5x', '-5x'].map((value) => messageFor({ top: value }))).toEqual(['x5', '5x', '-5x'].map(NOT_A_NUMBER));
  });
});

describe('a --skip', () => {
  it('is taken when it is a whole number, zero or more digits long', () => {
    expect([messageFor({ skip: '0' }), messageFor({ skip: '12' })]).toEqual([undefined, undefined]);
  });

  it('is refused when it is negative or not a whole number, with the reason', () => {
    expect(messageFor({ skip: '-3' })).toBe('must be a non-negative integer (Graph rejects skip=negatives)');
    expect(messageFor({ skip: '5x' })).toBe('must be a non-negative integer (got "5x", which is not a number)');
    expect(messageFor({ skip: 'x5' })).toBe('must be a non-negative integer (got "x5", which is not a number)');
  });
});

describe('the OData flags', () => {
  const names = (options: ReadonlyArray<{ readonly name: string }>): ReadonlyArray<string> => options.map((option) => option.name);

  it('are optional, keyed by their name, and each describes its OData parameter', () => {
    expect(odataQueryOptions.map(({ name, key, required }) => ({ name, key, required }))).toEqual(
      ['top', 'skip', 'select', 'filter', 'orderby', 'expand'].map((name) => ({ name, key: name, required: false }))
    );
    expect(odataQueryOptions.filter((option) => !option.description.startsWith(`OData $${option.name}: `))).toEqual([]);
  });

  it('come in the subsets the narrower builders advertise', () => {
    expect(names(selectExpandOptions)).toEqual(['select', 'expand']);
    expect(names(filterSelectOptions)).toEqual(['select', 'filter']);
    expect(names(noSkipOptions)).toEqual(['top', 'select', 'filter', 'orderby', 'expand']);
    expect(names(topOnlyOptions)).toEqual(['top']);
    expect(names(selectOnlyOptions)).toEqual(['select']);
    expect(names(pickODataOptions(['expand', 'top']))).toEqual(['top', 'expand']);
  });

  it('parse only the fields of their subset', () => {
    expect(Object.keys(noSkipShape)).toEqual(['top', 'select', 'filter', 'orderby', 'expand']);
    expect(Object.keys(topOnlyShape)).toEqual(['top']);
    expect(Object.keys(selectOnlyShape)).toEqual(['select']);
    expect(Object.keys(pickODataShape(['orderby', 'skip']))).toEqual(['orderby', 'skip']);
  });
});
