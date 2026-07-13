import { describe, expect, it } from 'bun:test';
import { parseRecipients } from './parse-recipients.ts';

describe('parseRecipients', () => {
  it('splits a comma-separated list into Graph address objects, trimming surrounding whitespace', () => {
    expect(parseRecipients('alice@example.com, bob@example.com')).toEqual([{ emailAddress: { address: 'alice@example.com' } }, { emailAddress: { address: 'bob@example.com' } }]);
  });

  it('drops empty segments from leading, trailing, and doubled commas', () => {
    expect(parseRecipients(', alice@example.com,, bob@example.com,')).toEqual([
      { emailAddress: { address: 'alice@example.com' } },
      { emailAddress: { address: 'bob@example.com' } },
    ]);
  });

  it('returns an empty list when the input carries no addresses', () => {
    expect(parseRecipients('  ,  ,')).toEqual([]);
  });
});
