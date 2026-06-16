import { describe, expect, it } from 'bun:test';
import { firstSentence } from './first-sentence.ts';

describe('firstSentence', () => {
  it('cuts at the first `. ` and keeps the trailing period', () => {
    expect(firstSentence('Fetch a thing. Then do more. And more.')).toBe('Fetch a thing.');
  });

  it('cuts at the first newline and drops the stray break', () => {
    expect(firstSentence('Fetch a thing\nsecond paragraph')).toBe('Fetch a thing');
  });

  it('returns the full text untouched when there is no sentence break', () => {
    expect(firstSentence('one short summary with no period-space or newline')).toBe('one short summary with no period-space or newline');
  });

  it('takes the EARLIER of a `. ` and a newline (period before newline)', () => {
    expect(firstSentence('Alpha. Beta\nGamma')).toBe('Alpha.');
  });

  it('takes the EARLIER of a `. ` and a newline (newline before period)', () => {
    expect(firstSentence('Alpha\nBeta. Gamma')).toBe('Alpha');
  });

  it('does not cut on a leading period-space (index 0 is ignored, not a real sentence break)', () => {
    // search(/\. /) would never be 0 here, but the >0 filter guards the newline-at-0 case too.
    expect(firstSentence('\nrest')).toBe('\nrest');
  });
});
