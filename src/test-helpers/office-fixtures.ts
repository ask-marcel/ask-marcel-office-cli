import {
  Bookmark,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
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
import { deflateSync } from 'node:zlib';
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

// A sheet with a fully-blank middle row — what Excel leaves behind when the used
// range is padded past the real data. Default `sheet_to_csv` emits it as a bare
// `,` line; the adapter drops it via `blankrows: false`.
const buildXlsxWithBlankRow = (): Uint8Array => {
  const sheet = XLSX.utils.aoa_to_sheet([['Name', 'Age'], [], ['Alice', 30]]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  return new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer);
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
    <definedName name="TrueHidden" hidden="true">'Very Secret'!$Z$9</definedName>
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
          // Comment id 1 anchored to a known body span (commentRangeStart/End in document.xml).
          new Paragraph({
            children: [
              new TextRun('Revenue was '),
              new CommentRangeStart(1),
              new TextRun('the Q4 revenue figure'),
              new CommentRangeEnd(1),
              new TextRun({ children: [new CommentReference(1)] }),
              new TextRun(' overall.'),
            ],
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

const ODF_META_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"';

const ODF_CONTENT_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"';

/**
 * An OpenDocument text fixture carrying the metadata `extractOdfMetadata`
 * surfaces: Dublin Core + ODF meta properties, a multi-keyword list, and two
 * user-defined custom fields. ODF is a ZIP, so the shared zip adapter reads it;
 * only `meta.xml` matters for metadata.
 */
const buildRichOdt = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text');
  zip.file(
    'meta.xml',
    `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${ODF_META_NS}><office:meta>` +
      '<meta:generator>LibreOffice/7.4.2</meta:generator>' +
      '<dc:title>Q4 Plan</dc:title>' +
      '<dc:creator>Vincent</dc:creator>' +
      '<dc:description>Internal draft</dc:description>' +
      '<meta:initial-creator>Alice</meta:initial-creator>' +
      '<meta:creation-date>2026-05-01T10:00:00</meta:creation-date>' +
      '<meta:editing-cycles>7</meta:editing-cycles>' +
      '<meta:keyword>budget</meta:keyword><meta:keyword>confidential</meta:keyword>' +
      '<meta:user-defined meta:name="ClientID">ACME-42</meta:user-defined>' +
      '<meta:user-defined meta:name="Reviewer">Bob</meta:user-defined>' +
      '</office:meta></office:document-meta>'
  );
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_CONTENT_NS}><office:body><office:text>` +
      '<text:h text:outline-level="1">Heading One</text:h>' +
      '<text:p>First <text:span>paragraph</text:span> body.</text:p>' +
      '<text:h text:outline-level="2">Sub heading</text:h>' +
      '<text:p>Spaced<text:s text:c="2"/>out<text:tab/>cell.</text:p>' +
      '<text:list><text:list-item><text:p>Item one</text:p></text:list-item>' +
      '<text:list-item><text:p>Item two</text:p><text:list><text:list-item><text:p>Nested item</text:p></text:list-item></text:list></text:list-item></text:list>' +
      '<table:table table:name="Table1">' +
      '<table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row>' +
      '<table:table-row><table:table-cell><text:p>A2</text:p></table:table-cell><table:table-cell><text:p>B2</text:p></table:table-cell></table:table-row>' +
      '</table:table>' +
      '<text:section text:display="none" text:name="Secret"><text:p>Hidden body text</text:p></text:section>' +
      '<text:p>Final paragraph.<office:annotation><dc:creator>Reviewer</dc:creator><dc:date>2026-06-01T09:00:00</dc:date><text:p>Check this number.</text:p></office:annotation></text:p>' +
      '</office:text></office:body></office:document-content>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/** An OpenDocument spreadsheet: two named sheets; the data sheet ends in a 16384-wide empty-cell tail to exercise the repeat cap + trailing trim. */
const buildRichOds = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_CONTENT_NS}><office:body><office:spreadsheet>` +
      '<table:table table:name="Budget">' +
      '<table:table-row><table:table-cell><text:p>Item</text:p></table:table-cell><table:table-cell><text:p>Cost</text:p></table:table-cell></table:table-row>' +
      '<table:table-row><table:table-cell><text:p>Rent</text:p></table:table-cell><table:table-cell><text:p>1000</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16384"/></table:table-row>' +
      '</table:table>' +
      '<table:table table:name="Notes"><table:table-row><table:table-cell><text:p>Hello</text:p></table:table-cell></table:table-row></table:table>' +
      '</office:spreadsheet></office:body></office:document-content>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/** An OpenDocument presentation: two named slides, each carrying a text box with paragraph text. */
const buildRichOdp = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation');
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_CONTENT_NS}><office:body><office:presentation>` +
      '<draw:page draw:name="Intro"><draw:frame><draw:text-box><text:p>Welcome slide</text:p><text:p>Subtitle here</text:p></draw:text-box></draw:frame></draw:page>' +
      '<draw:page draw:name="Details"><draw:frame><draw:text-box><text:p>Detail bullet</text:p></draw:text-box></draw:frame></draw:page>' +
      '</office:presentation></office:body></office:document-content>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/** A barebones ODF package: mimetype only, no meta.xml. */
const buildMinimalOdt = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text');
  return zip.generateAsync({ type: 'uint8array' });
};

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
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>' +
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

/**
 * A package carrying image media across all three format prefixes
 * (word/xl/ppt/media) — raster (png/jpeg/gif) plus an svg — alongside a legacy
 * vector part (.emf) and a non-media binary (embeddings) that the extractor
 * must skip. Exercises the image filter of the media extractor (svg included,
 * emf excluded). Hand-rolled so the byte contents are deterministic.
 */
const buildMediaSamples = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('word/media/image1.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
  zip.file('xl/media/photo.jpeg', new Uint8Array([0xff, 0xd8, 0xff, 0xe0]));
  zip.file('ppt/media/diagram.gif', new Uint8Array([0x47, 0x49, 0x46, 0x38]));
  zip.file('word/media/chart.svg', new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><text>Org chart label</text></svg>'));
  zip.file('ppt/media/logo.emf', new Uint8Array([0x01, 0x00, 0x00, 0x00]));
  zip.file('word/embeddings/oleObject1.bin', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
  return zip.generateAsync({ type: 'uint8array' });
};

/**
 * An adversarial deck built to exercise every branch of slide extraction:
 *  - slides added out of numeric order (slide10 before slide2) and two-digit, so
 *    numeric sort is distinguishable from list / lexical order;
 *  - slide10 puts a BODY placeholder before its title (slideTitle must skip
 *    non-title placeholders) and includes a whitespace-only paragraph (slideBodyText
 *    must filter blanks before joining);
 *  - slide2 is hidden and its notesSlide relationship is the SECOND rel, not the
 *    first (notesPathFor can't just grab rel #1), with multi-paragraph notes
 *    including a blank (notesText must filter, then join with a space).
 */
const buildAdversarialPptx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'ppt/slides/slide10.xml',
    `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS}><p:cSld><p:spTree>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>body first</a:t></a:r></a:p><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>The Title</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  zip.file(
    'ppt/slides/slide2.xml',
    `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS} show="0"><p:cSld><p:spTree>` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:sld>'
  );
  zip.file(
    'ppt/slides/_rels/slide2.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'ppt/notesSlides/notesSlide2.xml',
    `<?xml version="1.0"?><p:notes ${PPTX_SLIDE_NS}><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/>` +
      '<a:p><a:r><a:t>note one</a:t></a:r></a:p><a:p><a:r><a:t> </a:t></a:r></a:p><a:p><a:r><a:t>note two</a:t></a:r></a:p>' +
      '</p:txBody></p:sp></p:spTree></p:cSld></p:notes>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

/** A deck whose single slide has no shapes at all — empty title, body text, and notes. */
const buildEmptyPptx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld ${PPTX_SLIDE_NS}><p:cSld><p:spTree></p:spTree></p:cSld></p:sld>`);
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

const enc = (s: string): Buffer => Buffer.from(s, 'latin1');

// Minimal valid PDF (correct xref offsets) from a content stream, the page's
// Resources dict, and an optional object #5 (an image XObject or a font).
const buildPdf = (content: Buffer, resources: string, obj5?: Buffer): Uint8Array => {
  const objs: Record<number, Buffer> = {
    1: enc('<</Type/Catalog/Pages 2 0 R>>'),
    2: enc('<</Type/Pages/Kids[3 0 R]/Count 1>>'),
    3: enc(`<</Type/Page/Parent 2 0 R/MediaBox[0 0 100 100]/Contents 4 0 R/Resources${resources}>>`),
    4: Buffer.concat([enc(`<</Length ${content.length}>>\nstream\n`), content, enc('\nendstream')]),
  };
  if (obj5 !== undefined) objs[5] = obj5;
  const ids = Object.keys(objs)
    .map(Number)
    .sort((a, b) => a - b);
  let pdf = enc('%PDF-1.7\n');
  const off: Record<number, number> = {};
  for (const id of ids) {
    off[id] = pdf.length;
    pdf = Buffer.concat([pdf, enc(`${id} 0 obj\n`), objs[id]!, enc('\nendobj\n')]);
  }
  const xrefAt = pdf.length;
  const size = ids.length + 1;
  let xref = enc(`xref\n0 ${size}\n0000000000 65535 f \n`);
  for (const id of ids) xref = Buffer.concat([xref, enc(`${String(off[id]).padStart(10, '0')} 00000 n \n`)]);
  return new Uint8Array(Buffer.concat([pdf, xref, enc(`trailer\n<</Size ${size}/Root 1 0 R>>\nstartxref\n${xrefAt}\n%%EOF`)]));
};

const buildImageXObject = (): Buffer => {
  const img = deflateSync(Buffer.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]));
  return Buffer.concat([
    enc(`<</Type/XObject/Subtype/Image/Width 2/Height 2/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/FlateDecode/Length ${img.length}>>\nstream\n`),
    img,
    enc('\nendstream'),
  ]);
};

// `withImage` paints one 2x2 RGB image (FlateDecode); no-image paints an empty text object.
const buildPdfWithImage = (): Uint8Array => buildPdf(enc('q 50 0 0 50 25 25 cm /Im0 Do Q'), '<</XObject<</Im0 5 0 R>>>>', buildImageXObject());
const buildPdfNoImages = (): Uint8Array => buildPdf(enc('BT ET'), '<<>>');
// A born-digital PDF with a real text layer (Helvetica + a Tj string) — extractable by pdfjs/unpdf.
// (pdfjs drops the final glyph of this minimal no-/Widths font, so tests assert on a leading substring.)
const buildPdfWithText = (): Uint8Array =>
  buildPdf(enc('BT /F1 12 Tf 10 50 Td (Hello from the PDF) Tj ET'), '<</Font<</F1 5 0 R>>>>', enc('<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>'));

// A docx whose only text lives in surfaces mammoth drops: a header, a footer, and a
// text box (w:txbxContent). Hand-rolled raw OOXML because the `docx` lib can't emit text boxes.
const buildDocxWithHeaderFooterTextbox = async (): Promise<Uint8Array> => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Body paragraph.</w:t></w:r></w:p><w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>Callout box text</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p></w:body></w:document>`
  );
  zip.file('word/header1.xml', `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Confidential draft</w:t></w:r></w:p></w:hdr>`);
  zip.file('word/footer1.xml', `<?xml version="1.0"?><w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Page footer note</w:t></w:r></w:p></w:ftr>`);
  return zip.generateAsync({ type: 'uint8array' });
};

// A docx engineered to exercise every branch of the side-channel extractors with
// KNOWN values: a fully-attributed comment, tracked ins/del (plus an empty-text
// insertion that must be filtered), a hidden (w:vanish) run + an empty vanish run +
// a plain run with no w:rPr, a named bookmark + an empty-name bookmark (filtered),
// whitespace-padded + whitespace-only field codes and an empty w:fldSimple (filtered),
// whitespace-padded + whitespace-only text boxes, a two-digit header (header10 — the
// `\d+` quantifier), a whitespace-only header (filtered after trim), and two decoy
// parts that only match a regex with its ^/$ anchors removed.
const buildSideChannelDocx = async (): Promise<Uint8Array> => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  const body =
    '<w:bookmarkStart w:id="10" w:name="BM_named"/>' +
    '<w:bookmarkStart w:id="11" w:name=""/>' +
    '<w:ins w:id="20" w:author="InsAuthor" w:date="2026-01-01T00:00:00Z"><w:r><w:t>kept-ins</w:t></w:r></w:ins>' +
    '<w:ins w:id="21" w:author="EmptyIns" w:date="2026-01-02T00:00:00Z"><w:r></w:r></w:ins>' +
    '<w:del w:id="30" w:author="DelAuthor" w:date="2026-02-02T00:00:00Z"><w:r><w:delText>kept-del</w:delText></w:r></w:del>' +
    '<w:r><w:rPr><w:vanish/></w:rPr><w:t>secret-hidden</w:t></w:r>' +
    '<w:r><w:rPr><w:vanish/></w:rPr><w:t></w:t></w:r>' +
    '<w:r><w:t>visible</w:t></w:r>' +
    '<w:r><w:instrText>  MERGEFIELD Spaced  </w:instrText></w:r>' +
    '<w:r><w:instrText>   </w:instrText></w:r>' +
    '<w:fldSimple w:instr="  DOCVARIABLE FS  "><w:r><w:t>x</w:t></w:r></w:fldSimple>' +
    '<w:fldSimple w:instr=""><w:r><w:t>y</w:t></w:r></w:fldSimple>' +
    '<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>  box-text  </w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>' +
    '<w:p><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>   </w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>';
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  zip.file(
    'word/comments.xml',
    `<?xml version="1.0"?><w:comments xmlns:w="${W}"><w:comment w:id="5" w:author="Commenter" w:initials="CC" w:date="2026-03-03T00:00:00Z"><w:p><w:r><w:t>comment-body</w:t></w:r></w:p></w:comment></w:comments>`
  );
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>  HeaderOneProse  </w:t></w:r></w:p><w:p><w:r><w:instrText>PAGE</w:instrText></w:r></w:p></w:hdr>`
  );
  zip.file('word/header2.xml', `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>   </w:t></w:r></w:p></w:hdr>`);
  zip.file('word/header10.xml', `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>HeaderTenProse</w:t></w:r></w:p></w:hdr>`);
  zip.file('word/footer1.xml', `<?xml version="1.0"?><w:ftr xmlns:w="${W}"><w:p><w:r><w:t>FooterOneProse</w:t></w:r></w:p></w:ftr>`);
  zip.file('notword/header1.xml', `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>DECOY_NO_CARET</w:t></w:r></w:p></w:hdr>`);
  zip.file('word/header1.xmlbak', `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>DECOY_NO_DOLLAR</w:t></w:r></w:p></w:hdr>`);
  return zip.generateAsync({ type: 'uint8array' });
};

// A .zip carrying one of every entry kind the zip-conversion command branches on:
// each Office family (docx/xlsx/pptx/odt → markdown), a plain-text file (decoded
// inline), a malformed docx (conversion-failed note), and two non-convertible
// entries (a pdf + a raw binary → skip-note). Hand-rolled so the entry set is
// deterministic and exercises every dispatch branch.
const buildSampleZipArchive = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('report.docx', await buildSampleDocx());
  zip.file('data.xlsx', buildSampleXlsx());
  zip.file('deck.pptx', await buildRichPptx());
  zip.file('plan.odt', await buildRichOdt());
  zip.file('notes.txt', new TextEncoder().encode('hello from the archive'));
  zip.file('broken.docx', buildMalformedDocx());
  // A born-digital PDF → its text layer is extracted inline; a no-text PDF → skip note.
  zip.file('scan.pdf', buildPdfWithText());
  zip.file('blank.pdf', buildPdfNoImages());
  // Genuinely-binary payload (invalid UTF-8, no lead-byte continuation) → content-sniffs as non-text → skip note.
  zip.file('data.bin', new Uint8Array([0xff, 0xfe, 0xfd, 0x80]));
  // A dotless entry whose bytes ARE valid UTF-8 → the sniffer unpacks it as text (no extension needed).
  zip.file('LICENSE', new TextEncoder().encode('a dotless, no-extension entry'));
  // A dotless entry whose bytes are NOT valid UTF-8 → exercises the `ext === ''` skip-note branch ("no extension").
  zip.file('rawblob', new Uint8Array([0xff, 0xfe, 0xfd, 0x80]));
  return zip.generateAsync({ type: 'uint8array' });
};

// A .zip with more than the per-call entry cap (100), to exercise the truncation path.
const buildOversizedZipArchive = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  for (let i = 0; i < 101; i += 1) zip.file(`note${String(i).padStart(3, '0')}.txt`, new TextEncoder().encode(`entry ${i}`));
  return zip.generateAsync({ type: 'uint8array' });
};

// A minimal OpenDocument (.odt) carrying inline `xlink:href` hyperlinks in
// content.xml: one SharePoint link (twice, for dedup) and one non-SharePoint
// link (filtered out). The `mimetype` entry marks it as ODF so the command
// reads content.xml instead of relationship parts.
const buildOdtWithSharepointLinks = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text');
  zip.file(
    'content.xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">' +
      '<office:body><office:text>' +
      '<text:p><text:a xlink:href="https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/odf-spec.odt">spec</text:a></text:p>' +
      '<text:p><text:a xlink:href="https://example.com/not-sharepoint">other</text:a></text:p>' +
      '<text:p><text:a xlink:href="https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/odf-spec.odt">spec again</text:a></text:p>' +
      '</office:text></office:body></office:document-content>'
  );
  zip.file('styles.xml', '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>');
  return zip.generateAsync({ type: 'uint8array' });
};

// A minimal .docx whose relationship part carries external hyperlinks: one
// SharePoint link (twice, to exercise dedup), one non-SharePoint external link
// (must be filtered out), and one internal relationship (TargetMode is not
// External → must not be surfaced). Drives `extract-sharepoint-links-in-documents`.
const buildDocxWithSharepointLinks = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>');
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/spec.docx" TargetMode="External"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/not-sharepoint" TargetMode="External"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/spec.docx" TargetMode="External"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

export {
  buildAdversarialPptx,
  buildEmptyPptx,
  buildDocxWithHeaderFooterTextbox,
  buildDocxWithSharepointLinks,
  buildOdtWithSharepointLinks,
  buildSampleZipArchive,
  buildOversizedZipArchive,
  buildSideChannelDocx,
  buildMacroDocm,
  buildMalformedDocx,
  buildMalformedPptx,
  buildMalformedXlsx,
  buildMediaSamples,
  buildMinimalOdt,
  buildMinimalPptx,
  buildPdfNoImages,
  buildPdfWithImage,
  buildPdfWithText,
  buildRichDocx,
  buildRichOdp,
  buildRichOds,
  buildRichOdt,
  buildRichPptx,
  buildRichXlsx,
  buildSampleDocx,
  buildSampleXlsx,
  buildXlsxWithBlankRow,
};
