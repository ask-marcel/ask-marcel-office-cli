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

export { buildMalformedDocx, buildMalformedXlsx, buildRichDocx, buildRichXlsx, buildSampleDocx, buildSampleXlsx };
