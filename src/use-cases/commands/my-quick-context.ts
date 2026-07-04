import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

type MaybeId = { id?: string };

const valueOrUndefined = <T>(r: Result<unknown, GraphError>): T | undefined => (r.ok ? (r.value as T) : undefined);

// partial-result mode. The previous all-or-nothing
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

  const [meRes, driveRes, inboxRes, calendarRes, plannerRes, notebooksRes, teamsRes, recentRes, mailboxRes] = await Promise.all([
    graph.get('/me'),
    graph.get('/me/drive'),
    graph.get('/me/mailFolders/inbox'),
    graph.get('/me/calendar'),
    graph.get('/me/planner/plans?$top=1&$select=id,title'),
    graph.get('/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault'),
    graph.get('/me/joinedTeams?$top=1&$select=id,displayName'),
    graph.get('/me/drive/recent?$top=1&$select=id,name,lastModifiedDateTime'),
    // tenant timezone needed on first-contact so the
    // LLM stops treating every datetime as UTC. `mailboxSettings.timeZone` is
    // the closest thing Graph exposes (Outlook setting; matches the tz the
    // user sees in Outlook/Teams). $select keeps the payload small.
    graph.get('/me/mailboxSettings?$select=timeZone,language,workingHours'),
  ]);
  if (!meRes.ok) return meRes;

  // Audit follow-up (post-v1.3.0): the Microsoft To Do lists endpoint was
  // dropped from the parallel fan-out — it returned an array of
  // {id, displayName, wellknownListName} objects (typically 1-N entries
  // per user), which crowded the envelope with IDs an LLM rarely needs on
  // first contact. For Microsoft To Do discovery, call the dedicated
  // sibling command on demand; the absence here is the signal to do so.
  // Net effect: 10 → 9 parallel calls; same fan-out shape, slimmer payload.
  const me = meRes.value as { id?: string; displayName?: string; userPrincipalName?: string; mail?: string; jobTitle?: string };
  const drive = valueOrUndefined<MaybeId>(driveRes);
  const inbox = valueOrUndefined<MaybeId>(inboxRes);
  const calendar = valueOrUndefined<MaybeId>(calendarRes);
  const planner = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; title?: string }> }>(plannerRes);
  const notebooks = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; displayName?: string; isDefault?: boolean }> }>(notebooksRes);
  const teams = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; displayName?: string }> }>(teamsRes);
  const recent = valueOrUndefined<{ value?: ReadonlyArray<{ id?: string; name?: string; lastModifiedDateTime?: string }> }>(recentRes);
  const mailbox = valueOrUndefined<{ timeZone?: string; language?: { locale?: string }; workingHours?: { startTime?: string; endTime?: string; timeZone?: { name?: string } } }>(
    mailboxRes
  );

  return ok({
    // `jobTitle` added on user request — surfaces the user's role on first
    // contact so an LLM can answer "who am I working with" questions
    // without a second `get-current-user` call.
    user: { id: me.id, displayName: me.displayName, userPrincipalName: me.userPrincipalName, mail: me.mail, jobTitle: me.jobTitle },
    primaryDriveId: drive?.id,
    inboxId: inbox?.id,
    primaryCalendarId: calendar?.id,
    primaryPlannerPlanId: planner?.value?.[0]?.id,
    defaultNotebookId: notebooks?.value?.find((n) => n.isDefault === true)?.id ?? notebooks?.value?.[0]?.id,
    firstJoinedTeamId: teams?.value?.[0]?.id,
    recentDriveItemId: recent?.value?.[0]?.id,
    tenantTimeZone: mailbox?.timeZone,
    tenantLocale: mailbox?.language?.locale,
    tenantWorkingHours:
      mailbox?.workingHours?.startTime !== undefined && mailbox.workingHours.endTime !== undefined
        ? { start: mailbox.workingHours.startTime, end: mailbox.workingHours.endTime, timeZone: mailbox.workingHours.timeZone?.name }
        : undefined,
  });
};

const meta: CommandMeta = {
  summary:
    "One-shot discovery for the IDs every other command needs, plus the user's job title and tenant timezone / locale / working-hours. Issues 9 Graph calls in parallel and returns what each succeeded for. Partial-result mode: only `/me` is load-bearing — if any other sub-call fails (missing license, scope, or tenant policy) the corresponding field is `undefined` but the rest are still returned. Replaces the audit's 5-call discovery chain — feed the IDs straight into `list-mail-folder-messages`, `list-folder-files`, `list-planner-tasks`, `list-onenote-notebook-sections`, etc. For Microsoft To Do lists call `list-todo-task-lists` on demand (intentionally dropped from this command's fan-out — the array of {id, displayName, wellknownListName} entries crowded the envelope with IDs an LLM rarely needs on first contact). `tenantTimeZone` lets an LLM stop treating every datetime as UTC on first contact.",
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate:
    '(meta) parallel: /me, /me/drive, /me/mailFolders/inbox, /me/calendar, /me/planner/plans, /me/onenote/notebooks, /me/joinedTeams, /me/drive/recent, /me/mailboxSettings',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [],
  example: 'ask-marcel-office my-quick-context',
  responseShape:
    '`{ user: { id, displayName, userPrincipalName, mail, jobTitle? }, primaryDriveId?, inboxId?, primaryCalendarId?, primaryPlannerPlanId?, defaultNotebookId?, firstJoinedTeamId?, recentDriveItemId?, tenantTimeZone?, tenantLocale?, tenantWorkingHours?: { start, end, timeZone? } }` — every field except `user.id` is optional and absent when its source call failed. `user.jobTitle` is the user\'s role string from Azure AD (e.g. "Engineering Manager"). `tenantTimeZone` is the Outlook timezone string (e.g. "Romance Standard Time", "Pacific Standard Time"); `tenantLocale` is the IETF tag (e.g. "en-US"). For Microsoft To Do lists, call `list-todo-task-lists` separately — they were dropped from this command\'s fan-out to keep the envelope LLM-tractable.',
};

export { execute, meta, schema };
