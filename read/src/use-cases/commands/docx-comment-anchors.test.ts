import { describe, expect, it } from 'bun:test';
import { attr, clamp, extractCommentAnchors, leafText, tagOf } from './docx-comment-anchors.ts';

const doc = (body: string): string =>
  `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;

describe('extractCommentAnchors', () => {
  it('captures the run text between a commentRangeStart/End pair, ignoring text outside the range', () => {
    const xml = doc(
      '<w:p><w:r><w:t>Before </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>the span</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:t> after</w:t></w:r></w:p>'
    );
    const anchors = extractCommentAnchors(xml);
    expect(anchors.get('1')).toBe('the span');
    expect(anchors.size).toBe(1); // "Before"/"after" outside the range are not captured
  });

  it('concatenates multiple runs inside one range', () => {
    const xml = doc('<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>part one </w:t></w:r><w:r><w:t>part two</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>');
    expect(extractCommentAnchors(xml).get('2')).toBe('part one part two');
  });

  it('handles overlapping ranges — a run inside two open ranges is counted for both', () => {
    const xml = doc(
      '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>AAA</w:t></w:r><w:commentRangeStart w:id="4"/><w:r><w:t>BBB</w:t></w:r><w:commentRangeEnd w:id="4"/><w:r><w:t>CCC</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>'
    );
    const anchors = extractCommentAnchors(xml);
    expect(anchors.get('3')).toBe('AAABBBCCC');
    expect(anchors.get('4')).toBe('BBB');
  });

  it('captures text after an unclosed start marker (best-effort, no end)', () => {
    const xml = doc('<w:p><w:commentRangeStart w:id="9"/><w:r><w:t>tail</w:t></w:r></w:p>');
    expect(extractCommentAnchors(xml).get('9')).toBe('tail');
  });

  it('omits a range whose span is empty or whitespace-only', () => {
    const xml = doc('<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>   </w:t></w:r><w:commentRangeEnd w:id="7"/></w:p>');
    const anchors = extractCommentAnchors(xml);
    expect(anchors.has('7')).toBe(false);
    expect(anchors.size).toBe(0);
  });

  it('clamps a long span to 200 chars + ellipsis', () => {
    const long = 'x'.repeat(250);
    const xml = doc(`<w:p><w:commentRangeStart w:id="8"/><w:r><w:t>${long}</w:t></w:r><w:commentRangeEnd w:id="8"/></w:p>`);
    const span = extractCommentAnchors(xml).get('8');
    expect(span).toBe(`${'x'.repeat(200)}…`);
    expect(span?.length).toBe(201); // 200 chars + the ellipsis
  });

  it('returns an empty map for undefined or empty input', () => {
    expect(extractCommentAnchors(undefined).size).toBe(0);
    expect(extractCommentAnchors('').size).toBe(0);
  });

  it('returns an empty map when the document has no comment ranges', () => {
    expect(extractCommentAnchors(doc('<w:p><w:r><w:t>just body text</w:t></w:r></w:p>')).size).toBe(0);
  });
});

describe('docx-comment-anchors preserveOrder helpers', () => {
  it('tagOf returns the element tag, skipping the :@ attrs key (even when first) and #text leaves', () => {
    expect(tagOf({ ':@': { '@_w:id': '1' }, 'w:p': [] })).toBe('w:p');
    expect(tagOf({ 'w:r': [] })).toBe('w:r');
    expect(tagOf({ '#text': 'x' })).toBeUndefined();
    expect(tagOf({})).toBeUndefined();
  });

  it('attr reads a string @_<name> from the :@ bag, else empty string', () => {
    expect(attr({ ':@': { '@_w:id': '5' } }, 'w:id')).toBe('5');
    expect(attr({}, 'w:id')).toBe(''); // no :@ bag (undefined)
    expect(attr({ ':@': null }, 'w:id')).toBe(''); // null bag
    expect(attr({ ':@': { '@_w:id': 5 } }, 'w:id')).toBe(''); // non-string value
    expect(attr({ ':@': {} }, 'w:id')).toBe(''); // missing key
  });

  it('leafText concatenates only the string #text leaves', () => {
    expect(leafText([{ '#text': 'a' }, { '#text': 'b' }])).toBe('ab');
    expect(leafText([{ '#text': 'a' }, { 'w:noBreakHyphen': [] }])).toBe('a'); // non-leaf child contributes nothing
    expect(leafText([])).toBe('');
  });

  it('clamp trims, leaves ≤200 chars intact, and truncates 201+ to 200 chars + ellipsis', () => {
    expect(clamp('  hi  ')).toBe('hi');
    expect(clamp('x'.repeat(200))).toBe('x'.repeat(200)); // exactly 200 → NOT clamped (kills > vs >=)
    expect(clamp('x'.repeat(201))).toBe(`${'x'.repeat(200)}…`);
  });
});
