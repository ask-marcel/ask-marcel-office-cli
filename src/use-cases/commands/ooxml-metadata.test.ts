import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels, extractMacros } from './ooxml-metadata.ts';

const open = async (build: (zip: JSZip) => void): Promise<OoxmlZip> => {
  const zip = new JSZip();
  build(zip);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const opened = await openOoxmlZip(bytes);
  if (!opened.ok) throw new Error('fixture zip failed to open');
  return opened.value;
};

describe('extractCoreProps', () => {
  it('flattens cp:coreProperties children to a namespace-stripped record', async () => {
    const zip = await open((z) =>
      z.file(
        'docProps/core.xml',
        '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Vincent</dc:creator><dc:title>Q4</dc:title></cp:coreProperties>'
      )
    );
    expect(extractCoreProps(zip)).toEqual({ creator: 'Vincent', title: 'Q4' });
  });

  it('returns an empty record when docProps/core.xml is absent', async () => {
    const zip = await open((z) => z.file('[Content_Types].xml', '<x/>'));
    expect(extractCoreProps(zip)).toEqual({});
  });
});

describe('extractAppProps', () => {
  it('flattens extended properties, including nested vt:lpstr values', async () => {
    const zip = await open((z) =>
      z.file(
        'docProps/app.xml',
        '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Word</Application><Company>ACME</Company><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts></Properties>'
      )
    );
    expect(extractAppProps(zip)).toEqual({ Application: 'Microsoft Word', Company: 'ACME', TitlesOfParts: 'Sheet1' });
  });
});

describe('extractCustomProps', () => {
  it('reads each custom property name + first typed-value child', async () => {
    const zip = await open((z) =>
      z.file(
        'docProps/custom.xml',
        '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{x}" pid="2" name="ClientID"><vt:lpwstr>ACME-42</vt:lpwstr></property><property fmtid="{x}" pid="3" name="Reviewed"><vt:bool>true</vt:bool></property></Properties>'
      )
    );
    expect(extractCustomProps(zip)).toEqual([
      { name: 'ClientID', value: 'ACME-42' },
      { name: 'Reviewed', value: 'true' },
    ]);
  });

  it('returns an empty list when there are no custom properties', async () => {
    const zip = await open((z) => z.file('[Content_Types].xml', '<x/>'));
    expect(extractCustomProps(zip)).toEqual([]);
  });
});

describe('extractExternalRels', () => {
  it('surfaces only TargetMode="External" relationships, with the type tail and target, tagged by source part', async () => {
    const zip = await open((z) =>
      z.file(
        'word/_rels/document.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://evil.example/x" TargetMode="External"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
          '</Relationships>'
      )
    );
    expect(extractExternalRels(zip)).toEqual([{ source: 'word/_rels/document.xml.rels', type: 'hyperlink', target: 'https://evil.example/x' }]);
  });

  it('returns an empty list when no .rels part has an external target', async () => {
    const zip = await open((z) =>
      z.file(
        'word/_rels/document.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="x/image" Target="media/i.png"/></Relationships>'
      )
    );
    expect(extractExternalRels(zip)).toEqual([]);
  });
});

describe('extractMacros', () => {
  it('flags every vbaProject.bin part', async () => {
    const zip = await open((z) => {
      z.file('word/vbaProject.bin', new Uint8Array([0xd0, 0xcf]));
      z.file('word/document.xml', '<x/>');
    });
    expect(extractMacros(zip)).toEqual(['word/vbaProject.bin']);
  });

  it('returns an empty list for a macro-free package', async () => {
    const zip = await open((z) => z.file('word/document.xml', '<x/>'));
    expect(extractMacros(zip)).toEqual([]);
  });
});
