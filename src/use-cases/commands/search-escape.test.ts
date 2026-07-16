import { describe, expect, it } from 'bun:test';
import { kqlSearchClause, odataSearchQ } from './search-escape.ts';

describe('search-escape — wire-safe Graph search clause builders', () => {
  it('wraps a plain KQL query in the double quotes Graph requires and percent-encodes the value', () => {
    expect(kqlSearchClause('invoice')).toBe('$search=%22invoice%22');
  });

  it('escapes an embedded field:"phrase" quote pair so it reaches Graph as KQL phrase syntax (the live Contoso A2 & B7 case)', () => {
    expect(kqlSearchClause('subject:"Contoso A2 & B7 timeline"')).toBe('$search=%22subject%3A%5C%22Contoso%20R2%20%26%20B27%20timeline%5C%22%22');
  });

  it('turns a fully quoted phrase query into escaped KQL phrase quotes instead of a BadRequest', () => {
    expect(kqlSearchClause('"budget allocation"')).toBe('$search=%22%5C%22budget%20allocation%5C%22%22');
  });

  it('percent-encodes &, #, and + so they cannot truncate or corrupt the query string on the wire', () => {
    expect(kqlSearchClause('R&D #plan +next')).toBe('$search=%22R%26D%20%23plan%20%2Bnext%22');
  });

  it('doubles single quotes per OData string-literal escaping and percent-encodes for the search(q=) function form', () => {
    expect(odataSearchQ("john's plan & co")).toBe("john''s%20plan%20%26%20co");
  });

  it('passes a plain OData search(q=) value through with only percent-encoding', () => {
    expect(odataSearchQ('q1 budget')).toBe('q1%20budget');
  });
});
