import type { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandOptionMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

/**
 * `/me/drive/recent` and the three `/me/insights/*` listings name files without
 * the fields a triage needs (`lastModifiedDateTime`, `lastModifiedBy`), and
 * Graph ignores `$select` and `$expand=resource` on them (probed 2026-09-19).
 * `--with-item true` reads the driveItem behind every row after the listing,
 * one `get-drive-item` per row, and merges it as `item`; a row whose item
 * cannot be read carries `itemError` instead, so one broken link does not
 * fail the page.
 */

type Row = Record<string, unknown>;
type RowPath = (row: Row) => string | undefined;

const ITEM_SELECT = '$select=id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference';

const WITH_ITEM_OPTION: CommandOptionMeta = {
  name: 'with-item',
  key: 'withItem',
  required: false,
  description:
    'Pass `--with-item true` to read the driveItem behind every row after the listing (one get-drive-item per row) and merge it as `item`: `name`, `size`, `webUrl`, `lastModifiedDateTime`, `lastModifiedBy`, `createdBy`, `parentReference`. A row whose item cannot be read carries `itemError` instead of failing the page.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};

const nonEmpty = (v: unknown): v is string => typeof v === 'string' && v !== '';

/** An insight row points at its file through `resourceReference.id`, already shaped as `drives/{d}/items/{i}`. */
const insightItemPath: RowPath = (row) => {
  const id = (row['resourceReference'] as { readonly id?: unknown } | undefined)?.id;
  return nonEmpty(id) ? `/${id}` : undefined;
};

/** A recent-files row is the item itself when it is the user's own, or a `remoteItem` pointer when it lives in another drive. */
const recentItemPath: RowPath = (row) => {
  const remote = row['remoteItem'] as { readonly id?: unknown; readonly parentReference?: { readonly driveId?: unknown } } | undefined;
  if (nonEmpty(remote?.id) && nonEmpty(remote.parentReference?.driveId)) return `/drives/${remote.parentReference.driveId}/items/${remote.id}`;
  const driveId = (row['parentReference'] as { readonly driveId?: unknown } | undefined)?.driveId;
  return nonEmpty(row['id']) && nonEmpty(driveId) ? `/drives/${driveId}/items/${row['id']}` : undefined;
};

const enrichRow = async (graph: GraphClient, row: unknown, pathOf: RowPath): Promise<unknown> => {
  if (row === null || typeof row !== 'object') return row;
  const path = pathOf(row as Row);
  if (path === undefined) return { ...(row as Row), itemError: 'no drive item behind this row' };
  const item = await graph.get(`${path}?${ITEM_SELECT}`);
  return item.ok ? { ...(row as Row), item: item.value } : { ...(row as Row), itemError: item.error.message };
};

// A page of recent files is 200 rows by default (probed 2026-09-26) and the
// client does not retry a 429, so the item reads go out ten at a time, as
// `file-counts` does, rather than all at once.
const ITEM_READ_CHUNK = 10;

const enrichRows = async (graph: GraphClient, rows: ReadonlyArray<unknown>, pathOf: RowPath): Promise<ReadonlyArray<unknown>> => {
  const out: Array<unknown> = [];
  for (let start = 0; start < rows.length; start += ITEM_READ_CHUNK) {
    out.push(...(await Promise.all(rows.slice(start, start + ITEM_READ_CHUNK).map((row) => enrichRow(graph, row, pathOf)))));
  }
  return out;
};

/** Wraps a listing's execute: the merged schema validates the flag, the flag never reaches the OData schema, and the rows are enriched on request. */
const withItemEnrichment = (schema: z.ZodType, inner: Command['execute'], pathOf: RowPath): Command['execute'] => {
  const enriched: Command['execute'] = async (graph, params): Promise<Result<unknown, GraphError>> => {
    const parsed = schema.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const { withItem, ...rest } = params;
    const result = await inner(graph, rest);
    if (withItem !== 'true' || !result.ok) return result;
    const body = result.value as { readonly value?: unknown } & Row;
    if (!Array.isArray(body.value)) return result;
    return ok({ ...body, value: await enrichRows(graph, body.value, pathOf) });
  };
  return enriched;
};

export { insightItemPath, recentItemPath, WITH_ITEM_OPTION, withItemEnrichment };
