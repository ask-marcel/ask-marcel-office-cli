import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { readSheetsAsCsv } from '../../infra/sheetjs-adapter.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';

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
  const colCount = Math.max(...rows.map((r) => r.length));
  const padRow = (row: ReadonlyArray<string>): string => {
    const cells = Array.from({ length: colCount }, (_unused, i) => row[i] ?? '');
    return `| ${cells.join(' | ')} |`;
  };
  const header = padRow(rows[0] ?? []);
  const separator = `| ${Array.from({ length: colCount }, () => '---').join(' | ')} |`;
  const body = rows.slice(1).map((r) => padRow(r));
  return [header, separator, ...body].join('\n');
};

const xlsxToMarkdown = (bytes: Uint8Array): Result<MarkdownEnvelope, GraphError> => {
  const sheets = readSheetsAsCsv(bytes);
  if (!sheets.ok) return sheets;
  const sections = sheets.value.map(({ name, csv }) => {
    const table = csvToMarkdownTable(csv);
    return table.length === 0 ? `## ${name}` : `## ${name}\n\n${table}`;
  });
  const md = sections.join('\n\n');
  // size = UTF-8 byte count (audit §2.1); `md.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(md).byteLength, text: md });
};

export { csvToMarkdownTable, xlsxToMarkdown };
