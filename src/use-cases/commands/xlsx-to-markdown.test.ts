import { describe, expect, it } from 'bun:test';
import { buildMalformedXlsx, buildRichXlsx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { csvToMarkdownSection, csvToMarkdownTable, xlsxToMarkdown } from './xlsx-to-markdown.ts';

describe('xlsxToMarkdown', () => {
  it('converts an xlsx into one `## SheetName` section per sheet, each with a markdown table', async () => {
    const result = await xlsxToMarkdown(buildSampleXlsx());
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('text/markdown');
      expect(result.value.size).toBeGreaterThan(0);
      expect(result.value.text).toContain('## Sheet1');
      expect(result.value.text).toMatch(/\|\s*Name\s*\|\s*Age\s*\|\s*City\s*\|/);
      expect(result.value.text).toMatch(/\|\s*Alice\s*\|\s*30\s*\|\s*Paris\s*\|/);
      expect(result.value.text).toContain('## Sheet2');
      expect(result.value.text).toMatch(/\|\s*Product\s*\|\s*Price\s*\|/);
    }
  });

  it('separates the per-sheet sections with blank lines', async () => {
    const result = await xlsxToMarkdown(buildSampleXlsx());
    if (result.ok) {
      const sheet1Index = result.value.text.indexOf('## Sheet1');
      const sheet2Index = result.value.text.indexOf('## Sheet2');
      expect(sheet1Index).toBeLessThan(sheet2Index);
      expect(result.value.text.slice(sheet1Index, sheet2Index)).toContain('\n\n');
    }
  });

  it('propagates the api_error from the sheetjs adapter when the input is not a valid xlsx', async () => {
    const result = await xlsxToMarkdown(buildMalformedXlsx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('xlsx parse failed');
    }
  });

  it('appends a `## Workbook metadata` section when --include-metadata is set, surfacing authored content the value-rendered tables hide', async () => {
    const result = await xlsxToMarkdown(await buildRichXlsx(), { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text).toContain('## Workbook metadata');
    expect(result.value.text).toContain('### Defined names');
    expect(result.value.text).toContain('SecretFormula');
    expect(result.value.text).toContain('### Hidden / very-hidden sheets');
    expect(result.value.text).toContain('Very Secret');
  });

  it('does NOT append the metadata block by default (backward-compat — existing callers see the same envelope)', async () => {
    const result = await xlsxToMarkdown(buildSampleXlsx());
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.text).not.toContain('## Workbook metadata');
  });

  it('caps an oversized sheet to its header + a truncation hint instead of rendering the full table when --max-cells is exceeded', async () => {
    const result = await xlsxToMarkdown(buildSampleXlsx(), { maxCells: 4 });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text).toContain('## Sheet1');
    expect(result.value.text).toContain('## Sheet2');
    expect(result.value.text).not.toContain('Alice');
    expect(result.value.text).not.toContain('| --- |');
    expect(result.value.text).toContain('get-excel-range');
  });
});

describe('csvToMarkdownSection', () => {
  it('renders a sheet within the cap as its header followed by the full markdown table', () => {
    expect(csvToMarkdownSection('Budget', 'a,b\nc,d', 4)).toBe('## Budget\n\n| a | b |\n| --- | --- |\n| c | d |');
  });

  it('renders only the header for a sheet with no data rows', () => {
    expect(csvToMarkdownSection('Empty', '')).toBe('## Empty');
    expect(csvToMarkdownSection('Blank', '\n\n')).toBe('## Blank');
  });

  it('omits the table and points at the band-by-band Excel range commands once a sheet exceeds the cell cap (cells = rows × cols, not rows + cols)', () => {
    const section = csvToMarkdownSection('Dense', 'a,b,c,d\ne,f,g,h\ni,j,k,l\nm,n,o,p', 10);
    expect(section).toContain('## Dense');
    expect(section).not.toContain('| --- |');
    expect(section).not.toContain('| e | f');
    expect(section).toContain('16 cells'); // 4 rows × 4 cols — the hint quotes the product, not the sum
    expect(section).toContain('get-excel-used-range');
    expect(section).toContain('get-excel-range');
    expect(section).toContain('--max-cells');
  });

  it('counts columns from a single-line sheet that has no row breaks', () => {
    const section = csvToMarkdownSection('Wide', 'a,b,c,d,e', 2);
    expect(section).toContain('## Wide');
    expect(section).not.toContain('| --- |');
    expect(section).toContain('get-excel-range');
  });

  it('counts the empty trailing cell of a single-line sheet ending in a comma when estimating its width', () => {
    // 'a,b,' is 3 columns (the trailing comma opens an empty third cell); at cap 2 that tips it over.
    expect(csvToMarkdownSection('Trailing', 'a,b,', 2)).not.toContain('| --- |');
    expect(csvToMarkdownSection('Trailing', 'a,b,', 2)).toContain('3 cells');
  });

  it('renders a sheet sitting exactly on the cap and truncates only once it is strictly exceeded', () => {
    expect(csvToMarkdownSection('OnCap', 'a,b\nc,d', 4)).toContain('| --- |');
    expect(csvToMarkdownSection('OverCap', 'a,b\nc,d\ne,f', 4)).not.toContain('| --- |');
  });

  it('uses the default 50 000-cell cap and never materialises the full table for a sheet far over it', () => {
    const csv = Array.from({ length: 200_000 }, () => 'col1,col2,col3').join('\n');
    const section = csvToMarkdownSection('Big', csv);
    expect(section.startsWith('## Big')).toBe(true);
    expect(section).not.toContain('| --- |');
    expect(section.length).toBeLessThan(800);
  });

  it('renders a small sheet under the default cap when no explicit cap is given', () => {
    expect(csvToMarkdownSection('Small', 'a,b\nc,d')).toContain('| --- |');
  });
});

describe('csvToMarkdownTable', () => {
  it('renders a header row + separator + data rows from a basic CSV', () => {
    const md = csvToMarkdownTable('a,b\n1,2\n3,4');
    expect(md).toBe('| a | b |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |');
  });

  it('returns an empty string when the CSV has no non-empty lines', () => {
    expect(csvToMarkdownTable('')).toBe('');
    expect(csvToMarkdownTable('\n\n')).toBe('');
  });

  it('handles cells with embedded commas by respecting CSV quoting', () => {
    const md = csvToMarkdownTable('"a,b",c\n"1,2",3');
    expect(md).toContain('| a,b | c |');
    expect(md).toContain('| 1,2 | 3 |');
  });

  it('handles escaped quotes inside a quoted cell ("" → ")', () => {
    const md = csvToMarkdownTable('a,b\n"he said ""hi""",2');
    expect(md).toContain('| he said "hi" | 2 |');
  });

  it('pads short rows with empty cells so every row has the same column count', () => {
    const md = csvToMarkdownTable('a,b,c\n1\n2,3');
    expect(md).toContain('| 1 |  |  |');
    expect(md).toContain('| 2 | 3 |  |');
  });

  // 30 s ceiling: building a 1.05M-row table is sub-second normally, but Stryker's
  // instrumented sandbox slows it past Bun's default 5 s per-test timeout — without
  // this the mutation dry-run fails on a non-mutated baseline.
  it('renders a table from a sheet with more rows than the spread-argument limit without crashing (regression: Math.max(...rows) RangeError on a huge xlsx)', () => {
    const csv = Array.from({ length: 1_050_000 }, () => 'a,b').join('\n');
    const md = csvToMarkdownTable(csv);
    expect(md.startsWith('| a | b |\n| --- | --- |\n| a | b |')).toBe(true);
  }, 30_000);
});
