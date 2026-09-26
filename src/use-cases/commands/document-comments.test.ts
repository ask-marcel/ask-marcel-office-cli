import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildCommentedXlsx, buildMalformedDocx, buildMalformedPptx, buildMalformedXlsx, buildRichDocx, buildRichPptx } from '../../test-helpers/office-fixtures.ts';
import { cellRef, commentsFromBytes, mentionsIn, slideAnchor } from './document-comments.ts';

describe('who a comment @-mentions', () => {
  it('finds the known names written after an @, whatever their case, once each', () => {
    expect(mentionsIn('@jordan avery and @Robin Chen, then @Jordan Avery again', ['Robin Chen', 'Jordan Avery', 'Alex Kim'])).toEqual(['Robin Chen', 'Jordan Avery']);
  });

  it('ignores a name written without its @, and an empty name', () => {
    expect(mentionsIn('Robin Chen said @ nothing', ['Robin Chen', ''])).toEqual([]);
  });
});

describe('where a workbook comment sits', () => {
  it('writes the cell after its sheet, quoting a sheet name that is not a plain word', () => {
    expect(cellRef('Summary', 'C3')).toBe('Summary!C3');
    expect(cellRef('Q3 Plan', 'B2')).toBe("'Q3 Plan'!B2");
    expect(cellRef("Robin's", 'A1')).toBe("'Robin''s'!A1");
    expect(cellRef(undefined, 'B2')).toBe('B2');
  });
});

describe('the comments of an Office file as one list', () => {
  it('reads a workbook: notes and threaded comments with their sheet, the legacy copies of threaded comments dropped', async () => {
    const r = await commentsFromBytes(await buildCommentedXlsx(), 'plan.xlsx');
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value).toEqual({
      format: 'xlsx',
      count: 3,
      comments: [
        { author: 'Alex Kim', anchor: "'Q3 Plan'!B2", text: 'Check the Fabrikam line', mentions: [] },
        { author: 'Robin Chen', date: '2026-09-18T08:00:00Z', anchor: 'Summary!C3', text: '@Jordan Avery can you confirm the Q3 total?', mentions: ['Jordan Avery'] },
        { author: 'Jordan Avery', date: '2026-09-18T09:30:00Z', anchor: 'Summary!C3', text: 'Confirmed.', mentions: [] },
      ],
    });
  });

  it('reads a document: each comment with the text it is anchored to', async () => {
    const r = await commentsFromBytes(await buildRichDocx(), 'report.docm');
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value).toEqual({
      format: 'docx',
      count: 1,
      comments: [{ author: 'Jordan Avery', date: '2026-05-12T10:00:00.000Z', anchor: 'the Q4 revenue figure', text: 'Please double-check this figure.', mentions: [] }],
    });
  });

  it('reads a deck: legacy and modern comments, with the slide when the deck names it', async () => {
    const r = await commentsFromBytes(await buildRichPptx(), 'board.pptx');
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value).toEqual({
      format: 'pptx',
      count: 2,
      comments: [
        { author: 'Alice Smith', date: '2026-05-15T09:00:00Z', anchor: 'slide 1', text: 'Fix the revenue figure on this slide.', mentions: [] },
        { author: 'Bob Jones', date: '2026-05-16T11:00:00Z', text: 'Can we add a source for this number?', mentions: [] },
      ],
    });
  });

  it('passes a damaged file of each kind through as the reader answered it', async () => {
    for (const [bytes, name] of [
      [buildMalformedDocx(), 'a.docx'],
      [buildMalformedXlsx(), 'a.xlsx'],
      [buildMalformedPptx(), 'a.pptx'],
    ] as const) {
      const r = await commentsFromBytes(bytes, name);
      expect(r.ok).toBe(false);
    }
  });

  it('refuses a file that is not a Word, Excel or PowerPoint file, naming what it reads', async () => {
    const txt = await commentsFromBytes(new Uint8Array([1]), 'notes.txt');
    expect(txt.ok).toBe(false);
    if (!txt.ok) {
      expect(txt.error).toMatchObject({ type: 'api_error', status: 415, code: 'unsupported_format' });
      expect(txt.error.message).toBe(
        'list-document-comments reads Word, Excel and PowerPoint files (docx, xlsx, pptx and their macro-enabled and template variants); this file is a .txt.'
      );
    }
    const bare = await commentsFromBytes(new Uint8Array([1]), 'README');
    if (!bare.ok) expect(bare.error.message).toEndWith('this file has no extension.');
  });
});

describe('mentions of people who did not comment', () => {
  it('knows a document reader from its people registry even when that person left no comment', async () => {
    const zip = await JSZip.loadAsync(await buildRichDocx());
    zip.file(
      'word/people.xml',
      '<?xml version="1.0"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="Robin Chen"><w15:presenceInfo w15:providerId="AD" w15:userId="robin.chen@example.com"/></w15:person></w15:people>'
    );
    const commentsXml = (await zip.file('word/comments.xml')?.async('string')) ?? '';
    zip.file('word/comments.xml', commentsXml.replace('Please double-check this figure.', '@Robin Chen please double-check this figure.'));
    const r = await commentsFromBytes(await zip.generateAsync({ type: 'uint8array' }), 'report.docx');
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value.comments.map((c) => [c.author, c.mentions])).toEqual([['Jordan Avery', ['Robin Chen']]]);
  });
});

describe('where a deck comment sits', () => {
  it('names the slide by its number, and keeps any other part name as it came', () => {
    expect(slideAnchor('slide1.xml')).toBe('slide 1');
    expect(slideAnchor('slide12.xml')).toBe('slide 12');
    expect(slideAnchor('xslide1.xml')).toBe('xslide1.xml');
    expect(slideAnchor('slide1.xml.rels')).toBe('slide1.xml.rels');
    expect(slideAnchor(undefined)).toBeUndefined();
  });
});
