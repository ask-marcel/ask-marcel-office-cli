import { describe, expect, it } from 'bun:test';
import {
  attrOf,
  collectOrderedText,
  collectText,
  findAll,
  findAllTexts,
  orderedAttrOf,
  orderedElements,
  orderedSiblingGroups,
  parseXml,
  parseXmlOrdered,
  textOf,
} from './ooxml-xml-walker.ts';

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

// The ordered API exists for one reason the grouped parse cannot serve: knowing
// which element FOLLOWED which. These tests drive it directly rather than
// through a .docx, so each guard in it is pinned by an assertion that changes
// when the guard changes.

describe('parseXmlOrdered', () => {
  it('returns undefined for undefined or empty input, and a sibling array for real XML', () => {
    expect(parseXmlOrdered(undefined)).toBeUndefined();
    expect(parseXmlOrdered('')).toBeUndefined();
    expect(Array.isArray(parseXmlOrdered('<r><a/></r>'))).toBe(true);
  });
});

describe('orderedSiblingGroups', () => {
  // The whole point: the grouped parse renders this paragraph as
  // {del:[d1,d3], ins:[i2,i4]}, which cannot be told from del,del,ins,ins.
  it('keeps two interleaved tags in the order they appear, which the grouped parse cannot express', () => {
    const groups = orderedSiblingGroups(parseXmlOrdered('<w:p><w:del w:id="1"/><w:ins w:id="2"/><w:del w:id="3"/><w:ins w:id="4"/></w:p>'));
    const paragraph = groups.find((g) => g.some((e) => e.tag === 'w:del'));
    expect(paragraph?.map((e) => e.tag)).toEqual(['w:del', 'w:ins', 'w:del', 'w:ins']);
    expect(paragraph?.map((e) => orderedAttrOf(e.node, 'w:id'))).toEqual(['1', '2', '3', '4']);
  });

  it('reports elements only, dropping the text nodes that sit between them', () => {
    const groups = orderedSiblingGroups(parseXmlOrdered('<w:p>loose text<w:r/>more text</w:p>'));
    expect(groups.at(-1)?.map((e) => e.tag)).toEqual(['w:r']);
  });

  it('emits no group for a level that holds no elements at all', () => {
    // <w:p> holds only text, so its child level contributes nothing: two groups
    // (the root's and the body's), never an empty third.
    const groups = orderedSiblingGroups(parseXmlOrdered('<w:body><w:p>only text</w:p></w:body>'));
    expect(groups.map((g) => g.map((e) => e.tag))).toEqual([['w:body'], ['w:p']]);
  });

  it('returns nothing for a root that is not a sibling array', () => {
    expect(orderedSiblingGroups(undefined)).toEqual([]);
    expect(orderedSiblingGroups('not a tree')).toEqual([]);
  });
});

describe('orderedElements', () => {
  it('lists every element in document order, a parent before what it contains', () => {
    const elements = orderedElements(parseXmlOrdered('<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body>'));
    expect(elements.map((e) => e.tag)).toEqual(['w:body', 'w:p', 'w:r', 'w:t', 'w:sectPr']);
  });

  it('skips text nodes so a range marker is never confused with the text beside it', () => {
    const elements = orderedElements(parseXmlOrdered('<w:p>text<w:moveFromRangeStart w:name="m1"/></w:p>'));
    expect(elements.map((e) => e.tag)).toEqual(['w:p', 'w:moveFromRangeStart']);
  });

  it('returns nothing for a root that is not a sibling array', () => {
    expect(orderedElements(undefined)).toEqual([]);
  });
});

describe('orderedAttrOf', () => {
  it('reads an attribute out of the :@ bag the ordered shape puts it in', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:ins w:id="7" w:author="Robin Chen"/>'));
    expect(orderedAttrOf(element?.node ?? {}, 'w:author')).toBe('Robin Chen');
    expect(orderedAttrOf(element?.node ?? {}, 'w:id')).toBe('7');
  });

  it('returns an empty string for an element carrying no attributes at all', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:r/>'));
    expect(orderedAttrOf(element?.node ?? {}, 'w:author')).toBe('');
  });

  it('returns an empty string for an attribute the element does not carry', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:ins w:id="7"/>'));
    expect(orderedAttrOf(element?.node ?? {}, 'w:author')).toBe('');
  });
});

describe('collectOrderedText', () => {
  it('flattens every matching descendant into one string, however deeply nested', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:ins><w:r><w:t>one </w:t></w:r><w:r><w:t>two</w:t></w:r></w:ins>'));
    expect(collectOrderedText(element?.node ?? {}, 'w:t')).toBe('one two');
  });

  it('reads the deletion text tag rather than the insertion one when asked for it', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:del><w:r><w:delText>gone</w:delText></w:r></w:del>'));
    expect(collectOrderedText(element?.node ?? {}, 'w:delText')).toBe('gone');
    expect(collectOrderedText(element?.node ?? {}, 'w:t')).toBe('');
  });

  it('yields an empty string when the matching element holds elements instead of text', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:ins><w:t><w:noText/></w:t></w:ins>'));
    expect(collectOrderedText(element?.node ?? {}, 'w:t')).toBe('');
  });

  it('yields an empty string when nothing in the subtree matches', () => {
    const [element] = orderedElements(parseXmlOrdered('<w:ins><w:r><w:drawing/></w:r></w:ins>'));
    expect(collectOrderedText(element?.node ?? {}, 'w:t')).toBe('');
  });
});
