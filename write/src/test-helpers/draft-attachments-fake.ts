import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';

/**
 * A mailbox of messages with attachments behind a fake write graph, for the
 * tests of the commands that add and remove an attachment of a mail draft.
 * As Graph does, the list of the attachments gives a size 170 bytes larger
 * than the file, and a GET of one attachment with no `$select` gives its
 * `contentBytes`. `calls` logs each request as `<METHOD> <path>` (an upload as
 * `UPLOAD <attachments path>`). `fail` makes a request fail, `answer` makes it
 * give another value, and `onCall` runs before each request, so a test can
 * change the mailbox between the plan and the write.
 */
export type StoredAttachment = {
  readonly id: string;
  readonly name: string;
  readonly contentBytes: string;
  readonly contentType?: string;
  readonly isInline?: boolean;
  readonly type?: string;
};

export type StoredDraft = { readonly id: string; isDraft?: boolean; parentFolderId: string; readonly attachments: StoredAttachment[] };

export type DraftAttachmentsFake = {
  readonly graph: WriteGraph;
  readonly messages: Map<string, StoredDraft>;
  readonly calls: string[];
  readonly fail: Map<string, GraphError>;
  readonly answer: Map<string, unknown>;
  onCall: (call: string) => void;
  /** Upload answers with no attachment id, as an upload session that ends with no Location. */
  uploadWithoutId: boolean;
};

export const OUTBOX_ID = 'outbox-id';
export const DRAFTS_ID = 'drafts-id';

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

/** The size that the list of the attachments gives for a file: 170 bytes more than the file. */
export const listedSize = (attachment: StoredAttachment): number => Buffer.from(attachment.contentBytes, 'base64').byteLength + 170;

const metaOf = (attachment: StoredAttachment): Record<string, unknown> => ({
  '@odata.type': attachment.type ?? '#microsoft.graph.fileAttachment',
  id: attachment.id,
  name: attachment.name,
  size: listedSize(attachment),
  contentType: attachment.contentType ?? 'application/octet-stream',
  isInline: attachment.isInline ?? false,
});

const segmentsOf = (path: string): ReadonlyArray<string> => (path.split('?')[0] ?? '').split('/');

export const createDraftAttachmentsFake = (seed: ReadonlyArray<StoredDraft>): DraftAttachmentsFake => {
  const messages = new Map(seed.map((message) => [message.id, structuredClone(message)]));
  const calls: string[] = [];
  const fail = new Map<string, GraphError>();
  const answer = new Map<string, unknown>();
  let made = 0;
  const fake: DraftAttachmentsFake = { graph: fakeWriteGraph(), messages, calls, fail, answer, onCall: () => undefined, uploadWithoutId: false };
  const log = (call: string): GraphError | undefined => {
    fake.onCall(call);
    calls.push(call);
    return fail.get(call);
  };
  const read = (path: string): unknown => {
    const [, , , messageId = '', part, attachmentId] = segmentsOf(path);
    if (path.startsWith('/me/mailFolders/outbox')) return { id: OUTBOX_ID };
    const message = messages.get(messageId);
    if (message === undefined) return undefined;
    if (part === undefined) return { id: message.id, isDraft: message.isDraft, parentFolderId: message.parentFolderId };
    if (attachmentId === undefined) return { value: message.attachments.map(metaOf) };
    const attachment = message.attachments.find((item) => item.id === attachmentId);
    if (attachment === undefined) return undefined;
    return path.includes('?') ? metaOf(attachment) : { ...metaOf(attachment), contentBytes: attachment.contentBytes };
  };
  const graph = fakeWriteGraph({
    get: async (path) => {
      const failed = log(`GET ${path}`);
      if (failed !== undefined) return err(failed);
      if (answer.has(`GET ${path}`)) return ok(answer.get(`GET ${path}`));
      const value = read(path);
      return value === undefined ? err(notFound) : ok(structuredClone(value));
    },
    delete: async (path) => {
      const failed = log(`DELETE ${path}`);
      if (failed !== undefined) return err(failed);
      const [, , , messageId = '', , attachmentId] = segmentsOf(path);
      const attachments = messages.get(messageId)?.attachments ?? [];
      const at = attachments.findIndex((item) => item.id === attachmentId);
      if (at < 0) return err(notFound);
      attachments.splice(at, 1);
      return ok({ ok: true });
    },
    uploadAttachment: async (target, file) => {
      const messageId = target.kind === 'mail-draft' ? target.messageId : '';
      const failed = log(`UPLOAD /me/messages/${messageId}/attachments`);
      if (failed !== undefined) return err(failed);
      const message = messages.get(messageId);
      if (message === undefined) return err(notFound);
      made += 1;
      const id = `att-${made}`;
      message.attachments.push({
        id,
        name: file.name,
        contentBytes: Buffer.from(file.bytes).toString('base64'),
        ...(file.contentType === undefined ? {} : { contentType: file.contentType }),
      });
      return ok({ ...(fake.uploadWithoutId ? {} : { id }), name: file.name, size: file.bytes.byteLength });
    },
  });
  return Object.assign(fake, { graph });
};

/** An attachment of `text` (UTF-8), with the name `name`. */
export const textAttachment = (id: string, name: string, text: string, extra: Partial<StoredAttachment> = {}): StoredAttachment => ({
  id,
  name,
  contentBytes: Buffer.from(text).toString('base64'),
  ...extra,
});
