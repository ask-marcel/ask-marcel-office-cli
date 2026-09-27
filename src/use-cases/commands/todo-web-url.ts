/**
 * The link the To Do web app opens a task with, which Graph's `todoTask` does not
 * carry: `https://to-do.office.com/tasks/id/<id>/details`, the id exactly as Graph
 * gives it. Confirmed by hand 2026-09-27; the same link with the id's `=`
 * padding percent-encoded opened To Do without the task, so `=` stays raw and
 * only other unsafe characters are encoded.
 */
const TODO_WEB = 'https://to-do.office.com/tasks/id';

const todoWebUrl = (id: string): string => `${TODO_WEB}/${encodeURIComponent(id).replaceAll('%3D', '=')}/details`;

// A delta page's deleted task (`@removed`) opens nothing, so it gets no link.
const withWebUrl = (task: unknown): unknown => {
  const { id, '@removed': removed } = task as { readonly id?: unknown; readonly '@removed'?: unknown };
  return typeof id === 'string' && removed === undefined ? { ...(task as Record<string, unknown>), webUrl: todoWebUrl(id) } : task;
};

/** A task, or a page of tasks under `value[]`, with each task's `webUrl`. */
const withTaskLinks = (body: unknown): unknown => {
  const value = (body as { readonly value?: unknown }).value;
  return Array.isArray(value) ? { ...(body as Record<string, unknown>), value: value.map(withWebUrl) } : withWebUrl(body);
};

export { withTaskLinks };
