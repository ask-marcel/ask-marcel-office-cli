import { describe, expect, it } from 'bun:test';
import { buildMalformedXlsx, buildRichXlsx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { formatXlsxMetadata } from './xlsx-metadata-to-markdown.ts';
import { extractXlsxMetadata } from './xlsx-metadata.ts';

describe('extractXlsxMetadata', () => {
  it('surfaces every authored surface a value-rendered workbook hides — custom props, defined names, hidden sheets, cell comment, person-resolved threaded comment, external workbook link', async () => {
    const bytes = await buildRichXlsx();
    const result = await extractXlsxMetadata(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const m = result.value;

    expect(m.core.creator).toBe('Vincent Delacourt');
    expect(m.custom).toContainEqual({ name: 'ClientID', value: 'ACME-42' });

    // Defined names — both the plain named range and the hidden one, with their formulas
    expect(m.definedNames).toContainEqual({ name: 'TaxRate', refersTo: 'Summary!$A$1', hidden: false });
    expect(m.definedNames.some((d) => d.name === 'SecretFormula' && d.hidden && d.refersTo.includes('1.5'))).toBe(true);

    // Hidden + veryHidden sheets are flagged (the visible "Summary" sheet is NOT listed here)
    expect(m.hiddenSheets).toContainEqual({ name: 'Hidden Data', state: 'hidden' });
    expect(m.hiddenSheets).toContainEqual({ name: 'Very Secret', state: 'veryHidden' });
    expect(m.hiddenSheets.some((s) => s.name === 'Summary')).toBe(false);

    // Legacy cell comment — authorId resolved through the <authors> list
    expect(m.comments).toContainEqual({ cell: 'B2', author: 'Alice Smith', text: 'Double-check this total' });

    // Threaded comment — personId resolved to the display name via xl/persons/person.xml
    expect(m.threadedComments).toContainEqual({ cell: 'C3', author: 'Bob Jones', date: '2026-05-20T10:00:00Z', text: 'Needs review before sign-off' });

    // Persons registry
    expect(m.people.some((p) => p.displayName === 'Bob Jones' && p.userId === 'bob@contoso.com')).toBe(true);

    // External workbook link surfaces via the shared all-*.rels scan
    expect(m.externalRels.some((r) => r.source.includes('externalLink1.xml.rels') && r.target.includes('other-model.xlsx'))).toBe(true);
  });

  it('returns empty arrays for every list section on a barebones workbook (sheetjs output has no defined names, comments, hidden sheets, or custom props)', async () => {
    const bytes = buildSampleXlsx();
    const result = await extractXlsxMetadata(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.custom).toEqual([]);
    expect(result.value.definedNames).toEqual([]);
    expect(result.value.hiddenSheets).toEqual([]);
    expect(result.value.comments).toEqual([]);
    expect(result.value.threadedComments).toEqual([]);
    expect(result.value.people).toEqual([]);
    expect(result.value.externalRels).toEqual([]);
  });

  it('returns an api_error Result when the workbook zip is malformed', async () => {
    const result = await extractXlsxMetadata(buildMalformedXlsx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('ooxml zip parse failed');
    }
  });
});

describe('formatXlsxMetadata', () => {
  it('renders a populated `## Workbook metadata` block with every section heading and the authored content', async () => {
    const extracted = await extractXlsxMetadata(await buildRichXlsx());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    const md = formatXlsxMetadata(extracted.value);

    expect(md).toContain('## Workbook metadata');
    expect(md).toContain('### Core properties');
    expect(md).toContain('**creator**: Vincent Delacourt');
    expect(md).toContain('### Custom properties');
    expect(md).toMatch(/\|\s*ClientID\s*\|\s*ACME-42\s*\|/);
    expect(md).toContain('### Defined names');
    expect(md).toContain('TaxRate');
    expect(md).toContain('### Hidden / very-hidden sheets');
    expect(md).toContain('Very Secret');
    expect(md).toContain('### Cell comments');
    expect(md).toContain('Double-check this total');
    expect(md).toContain('### Threaded comments');
    expect(md).toContain('Bob Jones');
    expect(md).toContain('### People');
    expect(md).toContain('### External relationships');
  });

  it('emits `_(none)_` placeholders for every empty section on a barebones workbook so the output stays grep-stable', async () => {
    const extracted = await extractXlsxMetadata(buildSampleXlsx());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    const md = formatXlsxMetadata(extracted.value);
    expect(md).toContain('## Workbook metadata');
    expect(md).toContain('### Defined names\n\n_(none)_');
    expect(md).toContain('### Hidden / very-hidden sheets\n\n_(none)_');
    expect(md).toContain('### Cell comments\n\n_(none)_');
    expect(md).toContain('### Threaded comments\n\n_(none)_');
  });
});
