import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { extractXlsxMetadata } from './xlsx-metadata.ts';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string): string => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const note = (ref: string, authorId: string, text: string): string => `<comment ref="${ref}" authorId="${authorId}"><text><t>${text}</t></text></comment>`;
const legacyPart = (authors: ReadonlyArray<string>, notes: string): string => {
  const authorList = authors.map((a) => `<author>${a}</author>`).join('');
  return `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authorList}</authors><commentList>${notes}</commentList></comments>`;
};
const threadPart = (ref: string, text: string): string =>
  `<?xml version="1.0"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="${ref}" dT="2026-09-18T08:00:00Z" personId="{P1}" id="{T1}"><text>${text}</text></threadedComment></ThreadedComments>`;

// Two sheets; the parts map lets each case add its own comments and relationships.
const workbookWith = async (parts: Readonly<Record<string, string>>): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>'
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(`<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/>`)
  );
  for (const [path, body] of Object.entries(parts)) zip.file(path, body);
  return zip.generateAsync({ type: 'uint8array' });
};

const metadataOf = async (parts: Readonly<Record<string, string>>): Promise<Awaited<ReturnType<typeof extractXlsxMetadata>>> => extractXlsxMetadata(await workbookWith(parts));

describe('which sheet a workbook comment belongs to', () => {
  it('follows only comments relationships, so another relationship aimed at a comments part does not move it', async () => {
    const m = await metadataOf({
      'xl/comments1.xml': legacyPart(['Alex Kim'], note('B2', '0', 'Check the total')),
      'xl/threadedComments/threadedComment1.xml': threadPart('C3', 'Confirmed.'),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(
        `<Relationship Id="r1" Type="${REL}/comments" Target="../comments1.xml"/><Relationship Id="r2" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>`
      ),
      'xl/worksheets/_rels/sheet2.xml.rels': rels(
        `<Relationship Id="r1" Type="${REL}/drawing" Target="../comments1.xml"/><Relationship Id="r2" Type="${REL}/drawing" Target="../threadedComments/threadedComment1.xml"/>`
      ),
    });
    if (!m.ok) throw new Error(m.error.message);
    expect(m.value.comments.map((c) => c.sheet)).toEqual(['Summary']);
    expect(m.value.threadedComments.map((c) => c.sheet)).toEqual(['Summary']);
  });
});

describe('the author of a legacy note', () => {
  it('is the author at its index, and nobody when the index runs past the list', async () => {
    const m = await metadataOf({ 'xl/comments1.xml': legacyPart(['Alex Kim', 'Robin Chen'], note('A1', '1', 'second author') + note('A2', '7', 'unknown author')) });
    if (!m.ok) throw new Error(m.error.message);
    expect(m.value.comments.map((c) => [c.cell, c.author])).toEqual([
      ['A1', 'Robin Chen'],
      ['A2', ''],
    ]);
  });
});

describe('comments parts past the ninth', () => {
  it('are read like the first ones, legacy and threaded', async () => {
    const m = await metadataOf({
      'xl/comments10.xml': legacyPart(['Alex Kim'], note('D4', '0', 'tenth part')),
      'xl/threadedComments/threadedComment10.xml': threadPart('E5', 'tenth thread'),
    });
    if (!m.ok) throw new Error(m.error.message);
    expect(m.value.comments.map((c) => c.text)).toEqual(['tenth part']);
    expect(m.value.threadedComments.map((c) => c.text)).toEqual(['tenth thread']);
  });
});
