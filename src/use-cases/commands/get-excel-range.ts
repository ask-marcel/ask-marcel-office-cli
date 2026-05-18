import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';

// Graph happily expands an absurd range like `ZZ999999:AAA1` into ~2M
// "General" cells and streams ~76 MB back — enough to exhaust LLM context
// or self-DoS the host process. Cap the in-flight range at 100 000 cells
// (~5 MB JSON ceiling) at the schema level so the operator gets a clear
// error instead of a runaway response. Addresses without a `:` (single
// cells, named ranges) bypass the check — Graph will validate them.
const CELL_COUNT_CAP = 100_000;
const A1_CELL = /^\$?([A-Z]+)\$?(\d+)$/;

const colLettersToIndex = (letters: string): number => {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
};

const stripSheetPrefix = (address: string): string => {
  const bangIndex = address.indexOf('!');
  return bangIndex === -1 ? address : address.slice(bangIndex + 1);
};

const computeCellCount = (address: string): number | null => {
  const noSheet = stripSheetPrefix(address);
  const colonIndex = noSheet.indexOf(':');
  if (colonIndex === -1) return null;
  const left = noSheet.slice(0, colonIndex).toUpperCase();
  const right = noSheet.slice(colonIndex + 1).toUpperCase();
  const m1 = A1_CELL.exec(left);
  const m2 = A1_CELL.exec(right);
  if (m1 === null || m2 === null) return null;
  const c1 = colLettersToIndex(m1[1] ?? '');
  const c2 = colLettersToIndex(m2[1] ?? '');
  const r1 = Number.parseInt(m1[2] ?? '0', 10);
  const r2 = Number.parseInt(m2[2] ?? '0', 10);
  return (Math.abs(c2 - c1) + 1) * (Math.abs(r2 - r1) + 1);
};

const addressSchema = z
  .string()
  .min(1)
  .superRefine((value, ctx) => {
    const count = computeCellCount(value);
    if (count !== null && count > CELL_COUNT_CAP) {
      ctx.addIssue({
        code: 'custom',
        message: `spans ${count.toLocaleString()} cells (cap: ${CELL_COUNT_CAP.toLocaleString()}). A request this size would stream a multi-MB Graph response and likely exhaust LLM context. Split into narrower ranges (e.g. one column at a time, or row bands).`,
      });
    }
  });

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), worksheetId: z.string().min(1), address: addressSchema });
const inner = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets/${p.worksheetId}/range(address='${p.address}')`, schema);
const execute = wrapExcelExecute(inner.execute);

const meta: CommandMeta = {
  summary:
    'Get the cell values, formulas, and formats of a specific Excel range (e.g. `A1:C10`). The CLI caps the in-flight range at 100 000 cells to prevent runaway responses — split absurd ranges (`ZZ999999:AAA1` etc.) into smaller bands.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: "/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address='{address}')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-range',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    {
      name: 'worksheet-id',
      key: 'worksheetId',
      required: true,
      description: 'Worksheet ID or worksheet name. Returned by `ask-marcel list-excel-worksheets`.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'address',
      key: 'address',
      required: true,
      description:
        'A1-style range address, e.g. `A1:C10` or a single cell like `B7`. The CLI rejects ranges spanning more than 100 000 cells client-side to prevent runaway responses. Do NOT prefix with the worksheet name — `--worksheet-id` already pins the sheet, and a cross-sheet prefix like `OtherSheet!A1:C2` is rejected by Graph.',
      argumentHint: { kind: 'a1Address' },
    },
  ],
  example: "ask-marcel get-excel-range --drive-id 'b!1234' --item-id '01XLSX' --worksheet-id 'Sheet1' --address 'A1:C10'",
  responseShape: 'single Microsoft Graph `workbookRange` resource (values, formulas, format)',
};

export { execute, meta, schema };
