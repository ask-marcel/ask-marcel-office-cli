import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';

/**
 * An Outlook item attached to a mail, read the only way Graph allows. A plain
 * read of an itemAttachment carries no `item` (it comes with $expand only), and
 * an embedded MAIL's own attachments live nowhere but in its MIME source
 * (`$value`). So: expand, and for a mail fetch the source; an embedded event or
 * contact comes back as the expanded attachment.
 */
type EmbeddedItem = { readonly kind: 'mail'; readonly source: Uint8Array } | { readonly kind: 'other'; readonly attachment: Record<string, unknown> };

const readEmbeddedItem = async (graph: ReadGraph, attachmentPath: string): Promise<Result<EmbeddedItem, GraphError>> => {
  const expanded = await graph.get(`${attachmentPath}?$expand=microsoft.graph.itemattachment/item`);
  if (!expanded.ok) return expanded;
  const attachment = expanded.value as Record<string, unknown> & { readonly item?: Record<string, unknown> };
  if (attachment.item?.['@odata.type'] !== '#microsoft.graph.message') return ok({ kind: 'other', attachment });
  const source = await fetchRawBytes(graph, `${attachmentPath}/$value`);
  return source.ok ? ok({ kind: 'mail', source: source.value }) : source;
};

export { readEmbeddedItem };
export type { EmbeddedItem };
