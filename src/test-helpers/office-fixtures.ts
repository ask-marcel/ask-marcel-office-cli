import {
  Bookmark,
  DeletedTextRun,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  InsertedTextRun,
  Packer,
  Paragraph,
  SimpleMailMergeField,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

const TINY_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';

const tinyPngBytes = (): Uint8Array => {
  const binary = atob(TINY_PNG_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

const buildSampleDocx = async (): Promise<Uint8Array> => {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Sample Heading')] }),
          new Paragraph({
            children: [
              new TextRun('Hello '),
              new TextRun({ text: 'world', bold: true }),
              new TextRun(', this is '),
              new TextRun({ text: 'italic', italics: true }),
              new TextRun(' text.'),
            ],
          }),
          new Paragraph({
            children: [
              new ImageRun({
                type: 'png',
                data: tinyPngBytes(),
                transformation: { width: 100, height: 100 },
              }),
            ],
          }),
          new Table({
            rows: [
              new TableRow({
                children: [new TableCell({ children: [new Paragraph('A')] }), new TableCell({ children: [new Paragraph('B')] })],
              }),
              new TableRow({
                children: [new TableCell({ children: [new Paragraph('1')] }), new TableCell({ children: [new Paragraph('2')] })],
              }),
            ],
          }),
        ],
      },
    ],
  });
  const buffer = await Packer.toBuffer(doc);
  return new Uint8Array(buffer);
};

const buildSampleXlsx = (): Uint8Array => {
  const sheet1 = XLSX.utils.aoa_to_sheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'Paris'],
    ['Bob', 25, 'Berlin'],
  ]);
  const sheet2 = XLSX.utils.aoa_to_sheet([
    ['Product', 'Price'],
    ['Widget', 9.99],
    ['Gadget', 14.5],
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1');
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2');
  const arrayBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return new Uint8Array(arrayBuffer);
};

const buildMalformedDocx = (): Uint8Array => new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x00]);

const buildMalformedXlsx = (): Uint8Array => new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xff, 0xff, 0xff, 0xff]);

/**
 * A workbook fixture that embeds the side-channel surfaces sheetjs cannot
 * write (defined names, custom doc properties, legacy + threaded comments)
 * — hand-rolled as raw OOXML in a JSZip, the same pattern the docx
 * people.xml case needed. Carries: custom properties, defined names (one
 * visible, one hidden), a visible + hidden + veryHidden sheet, a legacy cell
 * comment with an author, a threaded comment whose personId resolves through
 * xl/persons/person.xml, and an external workbook link.
 */
const buildRichXlsx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>Vincent Delacourt</dc:creator>
  <dc:title>Budget Model</dc:title>
</cp:coreProperties>`
  );
  zip.file(
    'docProps/custom.xml',
    `<?xml version="1.0"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ClientID"><vt:lpwstr>ACME-42</vt:lpwstr></property>
</Properties>`
  );
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="Very Secret" sheetId="3" state="veryHidden" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="TaxRate">Summary!$A$1</definedName>
    <definedName name="SecretFormula" hidden="1">'Hidden Data'!$B$2*1.5</definedName>
  </definedNames>
</workbook>`
  );
  zip.file(
    'xl/comments1.xml',
    `<?xml version="1.0"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice Smith</author></authors>
  <commentList>
    <comment ref="B2" authorId="0"><text><r><t>Double-check this total</t></r></text></comment>
  </commentList>
</comments>`
  );
  zip.file(
    'xl/threadedComments/threadedComment1.xml',
    `<?xml version="1.0"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="C3" dT="2026-05-20T10:00:00Z" personId="{P1}" id="{T1}"><text>Needs review before sign-off</text></threadedComment>
</ThreadedComments>`
  );
  zip.file(
    'xl/persons/person.xml',
    `<?xml version="1.0"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="Bob Jones" id="{P1}" userId="bob@contoso.com" providerId="AD"/>
</personList>`
  );
  zip.file('xl/externalLinks/externalLink1.xml', '<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
  zip.file(
    'xl/externalLinks/_rels/externalLink1.xml.rels',
    `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///\\\\server\\share\\other-model.xlsx" TargetMode="External"/>
</Relationships>`
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/**
 * A docx fixture that embeds every "side-channel" surface the
 * --include-metadata flag is supposed to surface: core/app/custom doc
 * properties, a comment, a tracked insertion and deletion, a hidden-
 * formatted run (vanish), an external hyperlink (which serialises as a
 * HYPERLINK field), a MERGEFIELD via SimpleMailMergeField, and a bookmark.
 */
const buildRichDocx = async (): Promise<Uint8Array> => {
  const doc = new Document({
    creator: 'Vincent Delacourt',
    lastModifiedBy: 'Vincent Delacourt',
    title: 'Q4 Report',
    subject: 'Quarterly Numbers',
    description: 'Internal review draft',
    keywords: 'q4,finance,review',
    customProperties: [
      { name: 'ClientID', value: 'ACME-42' },
      { name: 'ReviewStatus', value: 'pending' },
    ],
    comments: {
      children: [
        {
          id: 1,
          author: 'Vincent Delacourt',
          initials: 'VD',
          date: new Date('2026-05-12T10:00:00Z'),
          children: [new Paragraph({ children: [new TextRun('Please double-check this figure.')] })],
        },
      ],
    },
    sections: [
      {
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Sample Heading')] }),
          new Paragraph({
            children: [new TextRun('Visible body. '), new TextRun({ text: 'This is hidden.', vanish: true }), new TextRun(' Resumed.')],
          }),
          new Paragraph({
            children: [
              new InsertedTextRun({ id: 100, author: 'Vincent Delacourt', date: '2026-05-12T10:05:00Z', children: [new TextRun({ text: 'inserted-phrase' })] }),
              new DeletedTextRun({ id: 101, author: 'Vincent Delacourt', date: '2026-05-12T10:06:00Z', text: 'deleted-phrase' }),
            ],
          }),
          new Paragraph({
            children: [new ExternalHyperlink({ link: 'https://example.com/secret-portal', children: [new TextRun({ text: 'portal link' })] })],
          }),
          new Paragraph({ children: [new SimpleMailMergeField('CustomerName')] }),
          new Paragraph({ children: [new Bookmark({ id: 'BM_intro', children: [new TextRun('introduction')] })] }),
        ],
      },
    ],
  });
  const buffer = await Packer.toBuffer(doc);
  return new Uint8Array(buffer);
};

const buildMalformedPptx = (): Uint8Array => new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00, 0x01]);

/**
 * A macro-enabled document (`.docm`) carrying a `word/vbaProject.bin` so the
 * macro-presence flag can be exercised. The bin content is the OLE/CFB magic
 * header — enough to be a realistic stand-in; we only detect presence, never
 * decompile. Hand-rolled because the `docx` package can't inject a VBA part.
 */
const buildMacroDocm = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>macro-enabled body</w:t></w:r></w:p></w:body></w:document>'
  );
  zip.file('word/vbaProject.bin', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
  return zip.generateAsync({ type: 'uint8array' });
};

const PPTX_SLIDE_NS =
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/**
 * A deck fixture carrying the authored-but-invisible content a slide PDF
 * never shows: custom properties, a slide tag, legacy + modern comment
 * authors and comments, an external hyperlink, a visible slide (title +
 * speaker notes) and a hidden slide (`show="0"`). Hand-rolled OOXML in a
 * JSZip — there is no pptx builder in the dependency tree.
 */
const buildRichPptx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'docProps/core.xml',
    '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Vincent Delacourt</dc:creator><dc:title>Board Deck</dc:title></cp:coreProperties>'
  );
  zip.file(
    'docProps/custom.xml',
    '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ClientID"><vt:lpwstr>ACME-42</vt:lpwstr></property></Properties>'
  );
  zip.file('ppt/tags/tag1.xml', `<?xml version="1.0"?><p:tagLst ${PPTX_SLIDE_NS}><p:tag name="REVIEW_STATE" val="confidential-draft"/></p:tagLst>`);
  zip.file(
    'ppt/commentAuthors.xml',
    `<?xml version="1.0"?><p:cmAuthorLst ${PPTX_SLIDE_NS}><p:cmAuthor id="0" name="Alice Smith" initials="AS" userId="alice@contoso.com" providerId="AD"/></p:cmAuthorLst>`
  );
  zip.file(
    'ppt/comments/comment1.xml',
    `<?xml version="1.0"?><p:cmLst ${PPTX_SLIDE_NS}><p:cm authorId="0" dt="2026-05-15T09:00:00Z" idx="1"><p:pos x="0" y="0"/><p:text>Fix the revenue figure on this slide.</p:text></p:cm></p:cmLst>`
  );
  zip.file(
    'ppt/authors.xml',
    '<?xml version="1.0"?><p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{B1}" name="Bob Jones" initials="BJ" userId="bob@contoso.com" providerId="AD"/></p188:authorLst>'
  );
  zip.file(
    'ppt/comments/modernComment1.xml',
    '<?xml version="1.0"?><p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm authorId="{B1}" created="2026-05-16T11:00:00Z"><p188:txBody><a:bodyPr/><a:p><a:r><a:t>Can we add a source for this number?</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>'
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS}><p:cSld><p:spTree>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly Review</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId2"/></a:rPr><a:t>see the portal</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/board-portal" TargetMode="External"/>' +
      '</Relationships>'
  );
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    `<?xml version="1.0"?><p:notes ${PPTX_SLIDE_NS}><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>Remember to mention the Q3 shortfall and the ACME contract renewal.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`
  );
  zip.file(
    'ppt/slides/slide2.xml',
    `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS} show="0"><p:cSld><p:spTree>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Internal Only — Do Not Present</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/** A barebones deck: one visible, untitled slide with no tags, comments, notes, or custom props. */
const buildMinimalPptx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>plain content</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
  );
  return zip.generateAsync({ type: 'uint8array' });
};

export {
  buildMacroDocm,
  buildMalformedDocx,
  buildMalformedPptx,
  buildMalformedXlsx,
  buildMinimalPptx,
  buildRichDocx,
  buildRichPptx,
  buildRichXlsx,
  buildSampleDocx,
  buildSampleXlsx,
};
