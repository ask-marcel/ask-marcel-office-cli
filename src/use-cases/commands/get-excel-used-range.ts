import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { mapWacError } from './excel-error.ts';
import { formatZodError } from './format-zod-error.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';

// Audit Alex-session §3 follow-up: the Graph `usedRange()` endpoint returns
// four parallel 2D arrays — `values`, `text`, `numberFormat`, `formulas`.
// `numberFormat` is sparse-by-default and dominated by the literal string
// "General" repeated cell-by-cell, blowing up the envelope (~125 KB for a
// 3×148 sheet was the audit observation). The CLI ships a slim default
// projection that keeps the structural fields + `values` (typically ~5-15 KB)
// and offers `--full true` to opt back into the raw four-array shape.
//
// `--max-cells` is the safety valve for `values[]` itself — Graph happily
// returns a million-cell `values` if `usedRange` spans the whole worksheet.
// Default cap 50 000 cells (~3-5 MB JSON); when exceeded, the slim response
// drops `values` and surfaces a hint pointing at `get-excel-range` for band-
// by-band reads. `--full true` ignores the cap (caller has opted into the
// full payload regardless of size).
const DEFAULT_MAX_CELLS = 50_000;

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  worksheetId: z.string().min(1),
  full: z.enum(['true', 'false']).optional(),
  maxCells: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

type WorkbookRange = {
  readonly address?: string;
  readonly rowCount?: number;
  readonly columnCount?: number;
  readonly values?: ReadonlyArray<ReadonlyArray<unknown>>;
  readonly text?: ReadonlyArray<ReadonlyArray<unknown>>;
  readonly numberFormat?: ReadonlyArray<ReadonlyArray<unknown>>;
  readonly formulas?: ReadonlyArray<ReadonlyArray<unknown>>;
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, worksheetId } = parsed.data;
  const fullMode = parsed.data.full === 'true';
  const maxCells = Number(parsed.data.maxCells ?? String(DEFAULT_MAX_CELLS));

  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/usedRange()`;
  const result = await graph.get(path);
  if (!result.ok) return err(mapWacError(result.error));
  const body = result.value as WorkbookRange;

  if (fullMode) return ok({ ...body, projection: 'full' });

  const rowCount = body.rowCount ?? 0;
  const columnCount = body.columnCount ?? 0;
  const cellCount = rowCount * columnCount;
  if (cellCount > maxCells) {
    return ok({
      address: body.address,
      rowCount,
      columnCount,
      projection: 'slim',
      truncated: true,
      maxCells,
      hint: `usedRange spans ${cellCount.toLocaleString()} cells (> --max-cells ${maxCells.toLocaleString()}). \`values[]\` omitted to keep the envelope small. Either raise \`--max-cells <N>\`, fetch specific bands via \`get-excel-range --address 'A1:Cn'\`, or pass \`--full true\` to bypass the cap and return all four arrays.`,
    });
  }

  return ok({
    address: body.address,
    rowCount,
    columnCount,
    values: body.values,
    projection: 'slim',
  });
};

const meta: CommandMeta = {
  summary:
    'Return the worksheet\'s used range — the bounding box of every non-empty cell — as a single Excel range. The CLI ships a slim default that strips the redundant `text` / `numberFormat` / `formulas` 2D arrays Graph returns (mostly `"General"` repeated cell-by-cell), keeping `address` / `rowCount` / `columnCount` / `values`. Pass `--full true` to return the raw four-array Graph shape. `--max-cells` (default 50 000) caps the size of the projected `values[]`; oversize ranges drop `values` and surface a hint pointing at `get-excel-range` for band-by-band reads. Avoids fetching the entire 1M × 16K-cell sheet when only a small data island is populated.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/usedRange()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-usedrange',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: DRIVE_ID_DESCRIPTION,
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the .xlsx file.',
    },
    {
      name: 'worksheet-id',
      key: 'worksheetId',
      required: true,
      description:
        "Accepts either the worksheet display name (e.g. `Sheet1`, `PROJECT`) or the worksheet `id` GUID returned by `list-excel-worksheets`. If neither matches Graph responds `itemNotFound: The requested resource doesn't exist.` — when that happens, double-check spelling case-sensitively against `ask-marcel list-excel-worksheets --drive-id <d> --item-id <i>`.",
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'full',
      key: 'full',
      required: false,
      description:
        'Pass `--full true` to return the raw Graph `workbookRange` shape with all four 2D arrays (`values`, `text`, `numberFormat`, `formulas`). Default (`--full false`, or omitted) drops the three redundant arrays and ships only `values`. The raw shape on a 3×148 sheet is ~125 KB (most of it duplicated `"General"` numberFormat strings); the slim default is ~5-15 KB.',
    },
    {
      name: 'max-cells',
      key: 'maxCells',
      required: false,
      description:
        'Cap (positive integer; default 50 000) on the size of the projected `values[]` in slim mode. When the used-range exceeds the cap, the response keeps `address` / `rowCount` / `columnCount` but drops `values[]` and adds `truncated: true` plus a hint pointing at `get-excel-range` for band-by-band reads. Ignored when `--full true` is set (the caller has opted into the full payload regardless of size).',
    },
  ],
  example: "ask-marcel get-excel-used-range --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'",
  responseShape:
    "Slim projection (default): `{ address, rowCount, columnCount, values, projection: 'slim' }` — `values[]` is the 2D cell-value array. Oversize variant: `{ address, rowCount, columnCount, projection: 'slim', truncated: true, maxCells, hint }` (no `values`). With `--full true`: the raw Graph `workbookRange` resource (adds `text`, `numberFormat`, `formulas` 2D arrays) plus `projection: 'full'`. Workbook Online (WAC) errors are translated to a clear `item is not an accessible Excel workbook` envelope (see `excel-error.ts`).",
};

export { execute, meta, schema };
