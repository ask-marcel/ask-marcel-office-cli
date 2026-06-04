import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { readSheetsAsCsv } from '../../infra/sheetjs-adapter.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { formatXlsxMetadata } from './xlsx-metadata-to-markdown.ts';
import { extractXlsxMetadata } from './xlsx-metadata.ts';

type XlsxToMarkdownOptions = { readonly includeMetadata?: boolean; readonly maxCells?: number };

// A dense sheet renders a markdown table proportional to rows × cols, so a 49 MB
// workbook with a genuinely large used range builds a multi-hundred-MB string and
// OOMs the process (exit 144). Mirroring `get-excel-used-range --max-cells`, any
// sheet whose estimated cell count exceeds this cap is emitted as a header + a
// one-line hint pointing at the band-by-band Excel range commands, never the full
// table — and the cell count is measured WITHOUT materialising the table.
const DEFAULT_MAX_CELLS = 50_000;

const splitCsvLine = (line: string): ReadonlyArray<string> => {
  const cells: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      cells.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  cells.push(current);
  return cells;
};

const csvToMarkdownTable = (csv: string): string => {
  const lines = csv.split('\n').filter((line) => line.length > 0);
  if (lines.length === 0) return '';
  const rows = lines.map((line) => splitCsvLine(line));
  // Reduce, not `Math.max(...rows.map(…))`: spreading a huge sheet's row-lengths
  // as call arguments overflows the engine's argument limit (RangeError: Maximum
  // call stack size exceeded) once a sheet exceeds ~1M rows.
  const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);
  const padRow = (row: ReadonlyArray<string>): string => {
    const cells = Array.from({ length: colCount }, (_unused, i) => row[i] ?? '');
    return `| ${cells.join(' | ')} |`;
  };
  const header = padRow(rows[0] ?? []);
  const separator = `| ${Array.from({ length: colCount }, () => '---').join(' | ')} |`;
  const body = rows.slice(1).map((r) => padRow(r));
  return [header, separator, ...body].join('\n');
};

// Newline count via `charCodeAt` (10 = `\n`) — an O(1)-memory scan that never
// allocates the per-row array `csvToMarkdownTable` builds, so an oversized sheet
// is detected before it can blow up the heap. `\n` is 0x0A in every encoding the
// sheetjs CSV writer emits.
const countNewlines = (csv: string): number => {
  let count = 0;
  for (let i = 0; i < csv.length; i += 1) if (csv.charCodeAt(i) === 10) count += 1;
  return count;
};

const truncationHint = (rows: number, columns: number, maxCells: number): string =>
  `> _Table omitted: this sheet's used range is ~${(rows * columns).toLocaleString()} cells (${rows.toLocaleString()} rows × ${columns.toLocaleString()} cols), over the \`--max-cells\` ${maxCells.toLocaleString()} render cap — rendering it would build a multi-hundred-MB string. Read it band-by-band: \`get-excel-used-range\` for the populated bounding box, then \`get-excel-range --address 'A1:Cn'\` per band — or raise the cap with \`--max-cells <N>\`._`;

// One `## SheetName` section per sheet, capping oversized sheets to a hint. Cell
// count is estimated cheaply (columns from the first row, rows from a newline
// scan) so an over-cap sheet is skipped before `csvToMarkdownTable` materialises it.
const csvToMarkdownSection = (name: string, csv: string, maxCells: number = DEFAULT_MAX_CELLS): string => {
  const newlineAt = csv.indexOf('\n');
  const firstLine = newlineAt === -1 ? csv : csv.slice(0, newlineAt);
  const columns = splitCsvLine(firstLine).length;
  const rows = countNewlines(csv) + 1;
  if (rows * columns > maxCells) return `## ${name}\n\n${truncationHint(rows, columns, maxCells)}`;
  const table = csvToMarkdownTable(csv);
  return table.length === 0 ? `## ${name}` : `## ${name}\n\n${table}`;
};

const xlsxToMarkdown = async (bytes: Uint8Array, opts: XlsxToMarkdownOptions = {}): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const sheets = readSheetsAsCsv(bytes);
  if (!sheets.ok) return sheets;
  const maxCells = opts.maxCells ?? DEFAULT_MAX_CELLS;
  const sections = sheets.value.map(({ name, csv }) => csvToMarkdownSection(name, csv, maxCells));
  let md = sections.join('\n\n');
  if (opts.includeMetadata === true) {
    const meta = await extractXlsxMetadata(bytes);
    if (!meta.ok) return meta;
    md = `${md}\n\n${formatXlsxMetadata(meta.value)}`;
  }
  // size = UTF-8 byte count (audit §2.1); `md.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(md).byteLength, text: md });
};

export { csvToMarkdownSection, csvToMarkdownTable, xlsxToMarkdown };
export type { XlsxToMarkdownOptions };
