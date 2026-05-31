import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { formatZodError } from './format-zod-error.ts';

/**
 * Renders an Excel chart to a PNG via Graph's chart `Image()` function. Unlike
 * the rest of the conversion surface, this endpoint returns the bytes as a
 * base64 string in the JSON `value` field (NOT a 302 redirect), so a plain
 * `graph.get` is enough. `width=0,height=0` = the chart's natural size;
 * `fittingMode='Fit'` preserves aspect ratio.
 */

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  worksheetId: z.string().min(1),
  chartId: z.string().min(1),
});

const base64ByteLength = (b64: string): number => Math.floor((b64.length * 3) / 4);

const run = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, worksheetId, chartId } = parsed.data;
  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetId}/charts/${chartId}/Image(width=0,height=0,fittingMode='Fit')`;
  const result = await graph.get(path);
  if (!result.ok) return result;
  const value = (result.value as { value?: unknown }).value;
  if (typeof value !== 'string' || value === '') {
    return err({ type: 'api_error', status: 502, message: 'Graph chart Image() returned no base64 PNG in the `value` field' });
  }
  return ok({ contentType: 'image/png', size: base64ByteLength(value), base64: value });
};

const execute = wrapExcelExecute(run);

const meta: CommandMeta = {
  summary:
    "Render a chart on an Excel worksheet as a PNG (base64). Calls Graph's chart `Image()` function (natural size, aspect-preserving) so a vision-capable LLM can read the plotted data itself — not just the chart's title / position metadata that `list-excel-worksheet-charts` returns. The chart id or name comes from `list-excel-worksheet-charts`.",
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: "/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts/{chart-id}/Image(width=0,height=0,fittingMode='Fit')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chart-image',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'OneDrive / SharePoint drive ID.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    {
      name: 'worksheet-id',
      key: 'worksheetId',
      required: true,
      description: 'Worksheet display name (e.g. `Sheet1`) or the worksheet `id` GUID returned by `list-excel-worksheets`.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'chart-id',
      key: 'chartId',
      required: true,
      description: 'Chart name (e.g. `Chart 1`) or `id`, as returned by `list-excel-worksheet-charts`.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: "ask-marcel get-excel-chart-image --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1' --chart-id 'Chart 1' --output-path ./chart.png",
  responseShape:
    '`{ contentType: "image/png", size, base64 }` — the rendered chart PNG, inlined. Pair with the global `--output-path <file>` to write the PNG to disk (the response then replaces `base64` with `savedTo`).',
  producesBytes: true,
};

export { execute, meta, schema };
