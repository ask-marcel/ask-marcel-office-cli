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

// Parse CSV into records, quote-aware ACROSS newlines — so a quoted cell with an
// embedded newline (Excel Alt+Enter) stays one cell instead of splitting the row.
// Fast path when there are no quotes at all: no field can contain a separator, so
// a plain comma/line split is exact AND allocation-light for the huge-sheet case.
const parseCsvRecords = (csv: string): string[][] => {
  if (!csv.includes('"')) return csv.split('\n').map((line) => (line.endsWith('\r') ? line.slice(0, -1) : line).split(','));
  const records: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let inQuotes = false;
  for (let i = 0; i < csv.length; i += 1) {
    const ch = csv[i];
    if (inQuotes) {
      if (ch === '"' && csv[i + 1] === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') inQuotes = false;
      else cell += ch;
    } else if (ch === '"') inQuotes = true;
    else if (ch === ',') {
      row.push(cell);
      cell = '';
    } else if (ch === '\n') {
      row.push(cell);
      records.push(row);
      row = [];
      cell = '';
    } else if (ch !== '\r') cell += ch;
  }
  row.push(cell);
  records.push(row);
  return records;
};

// Drop blank records (a lone empty cell from a blank line / trailing newline).
const parseNonEmptyRecords = (csv: string): string[][] => parseCsvRecords(csv).filter((r) => !(r.length === 1 && r[0] === ''));

// A markdown table cell cannot contain a raw `|` (splits the column) or a newline
// (breaks the row) — escape the pipe, fold embedded newlines to a single space.
const escapeCell = (cell: string): string => cell.replace(/[\r\n]+/g, ' ').replace(/\|/g, '\\|');

const renderTable = (records: ReadonlyArray<ReadonlyArray<string>>, colCount: number): string => {
  const padRow = (row: ReadonlyArray<string>): string => `| ${Array.from({ length: colCount }, (_unused, i) => escapeCell(row[i] ?? '')).join(' | ')} |`;
  const separator = `| ${Array.from({ length: colCount }, () => '---').join(' | ')} |`;
  return [padRow(records[0] ?? []), separator, ...records.slice(1).map(padRow)].join('\n');
};

const csvToMarkdownTable = (csv: string): string => {
  const records = parseNonEmptyRecords(csv);
  if (records.length === 0) return '';
  // Reduce, not `Math.max(...rows.map(…))`: spreading a huge sheet's row-lengths
  // overflows the engine's argument limit once a sheet exceeds ~1M rows.
  const colCount = records.reduce((max, r) => Math.max(max, r.length), 0);
  return renderTable(records, colCount);
};

const truncationHint = (rows: number, columns: number, maxCells: number): string =>
  `> _Table omitted: this sheet's used range is ~${(rows * columns).toLocaleString()} cells (${rows.toLocaleString()} rows × ${columns.toLocaleString()} cols), over the \`--max-cells\` ${maxCells.toLocaleString()} render cap — rendering it would build a multi-hundred-MB string. Read it band-by-band: \`get-excel-used-range\` for the populated bounding box, then \`get-excel-range --address 'A1:Cn'\` per band — or raise the cap with \`--max-cells <N>\`._`;

// Render a CSV to a markdown table, or a truncation hint when the cell count
// (rows × the widest row) exceeds `maxCells`. Parsing allocates ~O(input); the
// expensive O(rows×cols) table string is built only under the cap — so emitting
// the hint, not the table, is what protects against the multi-hundred-MB OOM.
const renderCsvCapped = (csv: string, maxCells: number = DEFAULT_MAX_CELLS): string => {
  const records = parseNonEmptyRecords(csv);
  if (records.length === 0) return '';
  const columns = records.reduce((max, r) => Math.max(max, r.length), 0);
  if (records.length * columns > maxCells) return truncationHint(records.length, columns, maxCells);
  return renderTable(records, columns);
};

// One `## SheetName` section per sheet, capping oversized sheets to a hint.
const csvToMarkdownSection = (name: string, csv: string, maxCells: number = DEFAULT_MAX_CELLS): string => {
  const body = renderCsvCapped(csv, maxCells);
  return body === '' ? `## ${name}` : `## ${name}\n\n${body}`;
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

export { csvToMarkdownSection, csvToMarkdownTable, renderCsvCapped, xlsxToMarkdown };
export type { XlsxToMarkdownOptions };
