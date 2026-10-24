import type { GraphError } from '../../infra/graph-client.ts';

// Graph has a known quirk where certain `--select` / `--orderby` values on the
// tasks endpoint return the opaque `RequestBroker--ParseUri: Invalid request`
// with no recovery hint — most reliably any combination that includes `title`.
// Three task commands hit the same endpoint, so the rewrite lives here once (a
// hardcoded copy in each had drifted — one carried a circular "fetch it per
// task via the single-task get" workaround that fails the same way).
const PARSE_URI_NEEDLE = 'RequestBroker--ParseUri';

type TodoQuirkParams = { readonly select?: string; readonly orderby?: string };

// Returns a friendlier error to surface in place of the opaque Graph one, or
// `undefined` when the error is not the title-quirk (caller keeps the original).
const rewriteTodoTitleQuirk = (error: GraphError, params: TodoQuirkParams): GraphError | undefined => {
  if (error.type !== 'api_error' || !error.message.includes(PARSE_URI_NEEDLE)) return undefined;
  if (params.select !== undefined)
    return {
      type: 'api_error',
      status: error.status,
      message: `Graph rejected --select=${params.select} on this tasks endpoint with RequestBroker--ParseUri (known quirk — some field combinations are unsupported, most reliably any combo that includes \`title\`). Drop \`title\` from --select and request the other fields, or call without --select and slim the response client-side.`,
      code: 'cli_rewrite_todo_select_title',
    };
  if (params.orderby !== undefined)
    return {
      type: 'api_error',
      status: error.status,
      message: `Graph rejected --orderby=${params.orderby} on this tasks endpoint with RequestBroker--ParseUri (known quirk — sorting on \`title\` is unsupported). Call this command without --orderby and sort the response client-side, or order by a numeric/date field like \`createdDateTime\` / \`importance\` instead.`,
      code: 'cli_rewrite_todo_orderby_title',
    };
  return undefined;
};

export { rewriteTodoTitleQuirk };
