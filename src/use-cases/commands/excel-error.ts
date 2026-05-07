import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

const WAC_NEEDLE = 'Could not obtain a WAC access token';

const mapWacError = (e: GraphError): GraphError => {
  if (e.type !== 'api_error') return e;
  if (!e.message.includes(WAC_NEEDLE)) return e;
  return {
    type: 'api_error',
    status: e.status,
    message:
      'item is not an accessible Excel workbook — Graph rejected the /workbook endpoint with "AccessDenied: Could not obtain a WAC access token". Common causes: the item is a folder, a non-.xlsx file, or has a sensitivity label that blocks Office Online. Verify with `get-drive-item --drive-id <id> --item-id <id>` and check `file.mimeType` is `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.',
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
