import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

type MaybeId = { id?: string };

const valueOrUndefined = <T>(r: Result<unknown, GraphError>): T | undefined => (r.ok ? (r.value as T) : undefined);

// Audit round-7 Wave H: partial-result mode. The previous all-or-nothing
// failure (any single Graph call failing aborts the whole command) made
// cold-start unusable in tenants with quirky scopes — e.g. a user with
// `/me/drive` available but no Planner license would never get past
// `/me/planner/plans` and would lose the IDs they could have gotten. Now
// each sub-call's result is optional and the response carries the IDs it
// managed to fetch. Only `/me` is load-bearing — if that fails the whole
// session is broken and we surface the error.
const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });

  const [meRes, driveRes, inboxRes, listsRes, calendarRes, plannerRes, notebooksRes, teamsRes, recentRes] = await Promise.all([
    graph.get('/me'),
    graph.get('/me/drive'),
    graph.get('/me/mailFolders/inbox'),
    graph.get('/me/todo/lists'),
    graph.get('/me/calendar'),
    graph.get('/me/planner/plans?$top=1&$select=id,title'),
    graph.get('/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault'),
    graph.get('/me/joinedTeams?$top=1&$select=id,displayName'),
    graph.get('/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime'),
  ]);
  if (!meRes.ok) return meRes;

  const me = meRes.value as { id?: string; displayName?: string; userPrincipalName?: string; mail?: string };
  const drive = valueOrUndefined<MaybeId>(driveRes);
  const inbox = valueOrUndefined<MaybeId>(inboxRes);
  const lists = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; displayName?: string; wellknownListName?: string }> }>(listsRes);
  const calendar = valueOrUndefined<MaybeId>(calendarRes);
  const planner = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; title?: string }> }>(plannerRes);
  const notebooks = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; displayName?: string; isDefault?: boolean }> }>(notebooksRes);
  const teams = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; displayName?: string }> }>(teamsRes);
  const recent = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; name?: string; lastModifiedDateTime?: string }> }>(recentRes);

  return ok({
    user: { id: me.id, displayName: me.displayName, userPrincipalName: me.userPrincipalName, mail: me.mail },
    primaryDriveId: drive?.id,
    inboxId: inbox?.id,
    todoLists: (lists?.value ?? []).map((l) => ({ id: l.id, displayName: l.displayName, wellknownListName: l.wellknownListName })),
    primaryCalendarId: calendar?.id,
    primaryPlannerPlanId: planner?.value?.[0]?.id,
    defaultNotebookId: notebooks?.value?.find((n) => n.isDefault === true)?.id ?? notebooks?.value?.[0]?.id,
    firstJoinedTeamId: teams?.value?.[0]?.id,
    recentDriveItemId: recent?.value?.[0]?.id,
  });
};

const meta: CommandMeta = {
  summary:
    "One-shot discovery for the IDs every other command needs. Issues nine Graph calls in parallel and returns the IDs each succeeded for. Partial-result mode: only `/me` is load-bearing — if any other sub-call fails (missing license, scope, or tenant policy) the corresponding field is `undefined` but the rest are still returned. Replaces the audit's 5-call discovery chain — feed the IDs straight into `list-mail-folder-messages`, `list-folder-files`, `list-todo-tasks`, `list-planner-tasks`, `list-onenote-notebook-sections`, etc.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate:
    '(meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/todo/lists, /me/calendar, /me/planner/plans, /me/onenote/notebooks, /me/joinedTeams, /me/drive/recent',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [],
  example: 'ask-marcel my-quick-context',
  responseShape:
    '{ user: { id, displayName, userPrincipalName, mail }, primaryDriveId?, inboxId?, todoLists: [{ id, displayName, wellknownListName }], primaryCalendarId?, primaryPlannerPlanId?, defaultNotebookId?, firstJoinedTeamId?, recentDriveItemId? } — every ID except `user.id` is optional and absent when its source call failed.',
};

export { execute, meta, schema };
