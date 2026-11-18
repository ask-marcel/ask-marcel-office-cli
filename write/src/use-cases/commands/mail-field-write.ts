import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import type { MailMessage, Written } from './mail-batch.ts';
import { isNotThere } from './mail-folder-reader.ts';

/*
 * The change of one state field of a message (`PATCH /me/messages/{id}`): the
 * read state, the follow-up flag or the categories. The run keeps the value of
 * before and the value it set; the undo writes the value of before back, only
 * if the field still has the value the run set. The undo compares this one
 * field, not the changeKey: to open a message changes its changeKey.
 */

const flagSchema = z.object({ flagStatus: z.string(), dueDateTime: z.object({ dateTime: z.string() }).nullish() }).loose();

const flagForm = (value: unknown): string => {
  const flag = flagSchema.safeParse(value);
  if (!flag.success) return 'notFlagged';
  const due = flag.data.dueDateTime?.dateTime.slice(0, 10);
  return due === undefined ? flag.data.flagStatus : `${flag.data.flagStatus}, due ${due}`;
};

// How the output shows each field, and how the undo compares it (two values with the same form are the same
// state), and the value of a field that Graph did not give, which the undo writes back.
const FIELDS = {
  isRead: { form: (value: unknown): string => String(value === true), empty: false },
  categories: {
    form: (value: unknown): string => (Array.isArray(value) ? value.map(String).toSorted((a, b) => a.localeCompare(b)) : []).join(', '),
    empty: [],
  },
  flag: { form: flagForm, empty: { flagStatus: 'notFlagged' } },
} as const;

/** A state field that a mail state command changes. */
export type MessageField = keyof typeof FIELDS;

/** The field as the output shows it, for example `true` or `flagged, due 2026-11-20`. */
export const fieldForm = (field: MessageField, value: unknown): string => FIELDS[field].form(value);

/** Writes `value` to the field of the message, with the undo that writes the value of before back. */
export const writeField = async (graph: WriteGraph, message: MailMessage, field: MessageField, value: unknown): Promise<Result<Written, GraphError>> => {
  const written = await graph.patch(`/me/messages/${message.id}`, { [field]: value });
  if (!written.ok) return written;
  return ok({
    after: { messageId: message.id, [field]: fieldForm(field, value) },
    undo: { kind: 'restore-message-field', data: { messageId: message.id, field, before: message.fields[field] ?? FIELDS[field].empty, set: value } },
  });
};

const fieldsSchema = z.record(z.string(), z.unknown());

const restoreSchema = z.object({ messageId: graphIdSchema, field: z.enum(['isRead', 'categories', 'flag']), before: z.unknown(), set: z.unknown() });

const notAnId: GraphError = {
  type: 'validation_error',
  message: 'The run journal names a message id that is not a Graph id, or a field that no command changes, so this undo does not run.',
};

/** The undo handler of a state field, by kind. */
export const FIELD_UNDO_HANDLERS: UndoHandlers = {
  'restore-message-field': {
    check: async (reads, data) => {
      const step = restoreSchema.safeParse(data);
      if (!step.success) return err(notAnId);
      const { messageId, field, set } = step.data;
      const read = await reads.get(`/me/messages/${messageId}?$select=${field}`);
      if (!read.ok) return isNotThere(read.error) ? ok('missing') : read;
      const now = fieldsSchema.safeParse(read.value).data?.[field];
      return ok(fieldForm(field, now) === fieldForm(field, set) ? 'ready' : 'changed');
    },
    apply: async (graph, data) => {
      const step = restoreSchema.safeParse(data);
      return step.success ? graph.patch(`/me/messages/${step.data.messageId}`, { [step.data.field]: step.data.before }) : err(notAnId);
    },
  },
};
