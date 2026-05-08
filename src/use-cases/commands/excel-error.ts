import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

const WAC_NEEDLE = 'Could not obtain a WAC access token';
// Graph's `/workbook` endpoint also surfaces a different leaky error when
// the underlying item is the wrong file type: a generic
// `InvalidRequest: Missing header Client-Request-Id. Header Client-Request-Id is not a guid.`
// (audit v1.0.0 §2.2). The header isn't actually missing — Graph reports
// that misleading error when it can't engage the Excel runtime for the
// item. Detect the same way and rewrite to the friendly envelope.
const NON_WORKBOOK_NEEDLES = [WAC_NEEDLE, 'Missing header Client-Request-Id'];

const FRIENDLY_NON_WORKBOOK =
  'item is not an accessible Excel workbook — Graph rejected the /workbook endpoint. Common causes: the item is a folder, a non-.xlsx file, or has a sensitivity label that blocks Office Online. Verify with `get-drive-item --drive-id <id> --item-id <id>` and check `file.mimeType` is `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.';

const mapWacError = (e: GraphError): GraphError => {
  if (e.type !== 'api_error') return e;
  if (!NON_WORKBOOK_NEEDLES.some((needle) => e.message.includes(needle))) return e;
  return {
    type: 'api_error',
    status: e.status,
    message: FRIENDLY_NON_WORKBOOK,
  };
};

type ExecuteFn = (graph: GraphClient, params: Record<string, string>) => Promise<Result<unknown, GraphError>>;

const wrapExcelExecute =
  (inner: ExecuteFn): ExecuteFn =>
  async (graph, params) => {
    const result = await inner(graph, params);
    if (result.ok) return result;
    return err(mapWacError(result.error));
  };

export { mapWacError, wrapExcelExecute };
