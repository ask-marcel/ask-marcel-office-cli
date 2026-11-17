import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';

/**
 * A small mailbox behind a fake write graph, for the dry-run and undo tests of
 * the draft commands. Each POST makes a draft `draft-<n>` (a create, a reply or
 * a forward); a PATCH merges its fields and gives a new changeKey; a DELETE
 * removes the message; a GET answers the message, or 404. `calls` logs each
 * request as `<METHOD> <path>`, and `edit` changes a message as Outlook would.
 */
export type Message = Record<string, unknown> & { readonly id: string; readonly isDraft: boolean; readonly changeKey: string };

export type MailboxFake = {
  readonly graph: WriteGraph;
  readonly messages: Map<string, Message>;
  readonly calls: string[];
  readonly edit: (id: string, fields: Record<string, unknown>) => void;
};

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

const idOf = (path: string): string => path.split('?')[0]?.split('/')[3] ?? '';

export const createMailboxFake = (seed: ReadonlyArray<Message> = []): MailboxFake => {
  const messages = new Map(seed.map((message) => [message.id, message]));
  const calls: string[] = [];
  let made = 0;
  const version = new Map<string, number>();
  const change = (id: string, fields: Record<string, unknown>): Result<unknown, GraphError> => {
    const message = messages.get(id);
    if (message === undefined) return err(notFound);
    const next = (version.get(id) ?? 1) + 1;
    version.set(id, next);
    const changed: Message = { ...message, ...fields, changeKey: `${id}-v${next}` };
    messages.set(id, changed);
    return ok(structuredClone(changed));
  };
  const graph = fakeWriteGraph({
    get: async (path) => {
      calls.push(`GET ${path}`);
      const message = messages.get(idOf(path));
      return message === undefined ? err(notFound) : ok(structuredClone(message));
    },
    post: async (path, body) => {
      calls.push(`POST ${path}`);
      made += 1;
      const id = `draft-${made}`;
      const message: Message = { ...(body as Record<string, unknown>), id, isDraft: true, changeKey: `${id}-v1` };
      messages.set(id, message);
      return ok(structuredClone(message));
    },
    patch: async (path, body) => {
      calls.push(`PATCH ${path}`);
      return change(idOf(path), body as Record<string, unknown>);
    },
    delete: async (path) => {
      calls.push(`DELETE ${path}`);
      messages.delete(idOf(path));
      return ok({ ok: true });
    },
  });
  return {
    graph,
    messages,
    calls,
    edit: (id, fields) => {
      change(id, fields);
    },
  };
};
