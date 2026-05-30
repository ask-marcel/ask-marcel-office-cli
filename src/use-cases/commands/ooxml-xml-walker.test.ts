import { describe, expect, it } from 'bun:test';
import { attrOf, collectText, findAll, findAllTexts, parseXml, textOf } from './ooxml-xml-walker.ts';

describe('parseXml', () => {
  it('returns undefined for undefined or empty input, and an object for real XML', () => {
    expect(parseXml(undefined)).toBeUndefined();
    expect(parseXml('')).toBeUndefined();
    expect(parseXml('<r><a:t>x</a:t></r>')).toBeTypeOf('object');
  });
});

describe('findAll', () => {
  it('returns a single matching element as a one-item list', () => {
    const matches = findAll(parseXml('<r><w:cm w:id="1"/></r>'), 'w:cm');
    expect(matches).toHaveLength(1);
    expect(attrOf(matches[0] ?? {}, 'w:id')).toBe('1');
  });

  it('returns every occurrence when the tag repeats (array shape)', () => {
    const matches = findAll(parseXml('<r><w:cm w:id="1"/><w:cm w:id="2"/></r>'), 'w:cm');
    expect(matches.map((m) => attrOf(m, 'w:id'))).toEqual(['1', '2']);
  });

  it('skips non-object members when a repeated tag mixes element and text nodes', () => {
    // `<x a="1"/>` parses to an object; `<x>t</x>` (no attrs) to a bare string.
    const matches = findAll(parseXml('<r><x a="1"/><x>t</x></r>'), 'x');
    expect(matches).toHaveLength(1);
    expect(attrOf(matches[0] ?? {}, 'a')).toBe('1');
  });

  it('returns an empty list when nothing matches', () => {
    expect(findAll(parseXml('<r><other/></r>'), 'w:cm')).toEqual([]);
  });
});

describe('textOf', () => {
  it('returns a string node verbatim', () => {
    expect(textOf('hello')).toBe('hello');
  });

  it('reads the #text child of an attributed element', () => {
    expect(textOf({ '@_x': '1', '#text': 'body' })).toBe('body');
  });

  it('returns empty string for an element with no #text, and for non-string/non-object nodes', () => {
    expect(textOf({ '@_x': '1' })).toBe('');
    expect(textOf(42)).toBe('');
    expect(textOf(undefined)).toBe('');
    expect(textOf(null)).toBe('');
  });
});

describe('attrOf', () => {
  it('returns a present string attribute', () => {
    expect(attrOf({ '@_w:id': '5' }, 'w:id')).toBe('5');
  });

  it('returns empty string for a missing attribute or a non-string value', () => {
    expect(attrOf({}, 'w:id')).toBe('');
    expect(attrOf({ '@_n': 5 }, 'n')).toBe('');
  });
});

describe('findAllTexts', () => {
  it('yields one entry per occurrence (single and repeated)', () => {
    expect(findAllTexts(parseXml('<r><w:instrText>HYPERLINK</w:instrText></r>'), 'w:instrText')).toEqual(['HYPERLINK']);
    expect(findAllTexts(parseXml('<r><w:instrText>A</w:instrText><w:instrText>B</w:instrText></r>'), 'w:instrText')).toEqual(['A', 'B']);
  });

  it('reads #text when the repeated leaf carries attributes', () => {
    expect(findAllTexts(parseXml('<r><w:t x="1">a</w:t><w:t x="2">b</w:t></r>'), 'w:t')).toEqual(['a', 'b']);
  });
});

describe('collectText', () => {
  it('concatenates the text of every matching descendant in document order', () => {
    expect(collectText(parseXml('<w:p><w:r><w:t>one </w:t><w:t>two</w:t></w:r></w:p>'), 'w:t')).toBe('one two');
  });

  it('reads a single descendant and returns empty string when none match', () => {
    expect(collectText(parseXml('<w:p><w:t>solo</w:t></w:p>'), 'w:t')).toBe('solo');
    expect(collectText(parseXml('<w:p><w:r/></w:p>'), 'w:t')).toBe('');
  });
});
