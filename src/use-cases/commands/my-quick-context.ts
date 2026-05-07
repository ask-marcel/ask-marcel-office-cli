import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

type MaybeId = { id?: string };

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });

  const [meRes, driveRes, inboxRes, listsRes, calendarRes] = await Promise.all([
    graph.get('/me'),
    graph.get('/me/drive'),
    graph.get('/me/mailFolders/inbox'),
    graph.get('/me/todo/lists'),
    graph.get('/me/calendar'),
  ]);
  if (!meRes.ok) return meRes;
  if (!driveRes.ok) return driveRes;
  if (!inboxRes.ok) return inboxRes;
  if (!listsRes.ok) return listsRes;
  if (!calendarRes.ok) return calendarRes;

  const me = meRes.value as { id?: string; displayName?: string; userPrincipalName?: string; mail?: string };
  const drive = driveRes.value as MaybeId;
  const inbox = inboxRes.value as MaybeId;
  const lists = listsRes.value as { value?: ReadonlyArray<{ id?: string; displayName?: string; wellknownListName?: string }> };
  const calendar = calendarRes.value as MaybeId;

  return ok({
    user: { id: me.id, displayName: me.displayName, userPrincipalName: me.userPrincipalName, mail: me.mail },
    primaryDriveId: drive.id,
    inboxId: inbox.id,
    todoLists: (lists.value ?? []).map((l) => ({ id: l.id, displayName: l.displayName, wellknownListName: l.wellknownListName })),
    primaryCalendarId: calendar.id,
  });
};

const meta: CommandMeta = {
  summary:
    "One-shot discovery for the IDs every other command needs. Issues five Graph calls in parallel (`/me`, `/me/drive`, `/me/mailFolders/inbox`, `/me/todo/lists`, `/me/calendar`) and returns `{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }`. Replaces the audit's 5-call discovery chain — feed the IDs straight into `list-mail-folder-messages`, `list-folder-files`, `list-todo-tasks`, etc.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [],
  example: 'ask-marcel my-quick-context',
  responseShape: '{ user: { id, displayName, userPrincipalName, mail }, primaryDriveId, inboxId, todoLists: [{ id, displayName, wellknownListName }], primaryCalendarId }',
};

export { execute, meta, schema };
