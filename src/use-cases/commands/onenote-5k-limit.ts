import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Audit round-8 H2: when a tenant exceeds the 5000-item OneNote limit per
 * document library, every OneNote read from a SharePoint site returns
 * Graph error code `10008` with a wall-of-text message pointing at a
 * MSDN blog. Wrap that with a one-line actionable summary so an LLM
 * doesn't surface 1+ KB of stale-blog prose to the user.
 */

const ONENOTE_5K_NEEDLE = '10008';

const wrapOnenote5kLimit =
  (inner: Command['execute']): Command['execute'] =>
  async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
    const result = await inner(graph, params);
    if (result.ok) return result;
    if (result.error.type === 'api_error' && (result.error.code === ONENOTE_5K_NEEDLE || result.error.message.includes('10008'))) {
      return err({
        type: 'api_error',
        status: result.error.status,
        message:
          'OneNote read blocked: tenant has hit the 5000-item limit per document library (Graph error 10008). This is a SharePoint-side configuration issue, not a CLI bug — raise with the M365 admin to split the library or archive content. The CLI cannot work around it.',
        code: 'cli_rewrite_onenote_5k_limit',
      });
    }
    return result;
  };

export { wrapOnenote5kLimit };
