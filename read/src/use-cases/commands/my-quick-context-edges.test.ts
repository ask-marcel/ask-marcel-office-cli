import { describe, expect, it } from 'bun:test';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { execute } from './my-quick-context.ts';

const FAILED: GraphError = { type: 'api_error', status: 403, message: 'denied' };
const ME = { id: 'u1', displayName: 'Alice' };
const NOTEBOOKS = '/me/onenote/notebooks?$top=1&$select=id,displayName,isDefault';
const MAILBOX = '/me/mailboxSettings?$select=timeZone,language,workingHours';

// Each route answers from `responses`; a route that is not there fails.
const graphWith = (responses: Record<string, Result<unknown, GraphError>>): GraphClient =>
  fakeGraphClient({ get: async (path) => (Object.hasOwn(responses, path) ? responses[path] : err(FAILED)) });

const contextWith = async (responses: Record<string, Result<unknown, GraphError>>): Promise<Record<string, unknown>> => {
  const result = await execute(graphWith({ '/me': ok(ME), ...responses }), {});
  if (!result.ok) throw new Error(`expected ok, got ${result.error.message}`);
  return result.value as Record<string, unknown>;
};

describe('my-quick-context at its edges', () => {
  it('refuses an unknown option before any call', async () => {
    const result = await execute(graphWith({}), { extra: '1' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('keeps only the user when every optional sub-call fails', async () => {
    expect(await contextWith({})).toEqual({
      user: { id: 'u1', displayName: 'Alice', userPrincipalName: undefined, mail: undefined, jobTitle: undefined },
      primaryDriveId: undefined,
      inboxId: undefined,
      primaryCalendarId: undefined,
      primaryPlannerPlanId: undefined,
      defaultNotebookId: undefined,
      firstJoinedTeamId: undefined,
      recentDriveItemId: undefined,
      tenantTimeZone: undefined,
      tenantLocale: undefined,
      tenantWorkingHours: undefined,
    });
  });

  it('gives no id for an empty listing', async () => {
    const context = await contextWith({ '/me/planner/plans?$top=1&$select=id,title': ok({ value: [] }), [NOTEBOOKS]: ok({ value: [] }) });
    expect(context.primaryPlannerPlanId).toBeUndefined();
    expect(context.defaultNotebookId).toBeUndefined();
  });

  it('picks the default notebook over the first one', async () => {
    const notebooks = ok({
      value: [
        { id: 'nb-a', isDefault: false },
        { id: 'nb-b', isDefault: true },
      ],
    });
    expect((await contextWith({ [NOTEBOOKS]: notebooks })).defaultNotebookId).toBe('nb-b');
  });

  it('falls back to the first notebook when no notebook is the default', async () => {
    expect((await contextWith({ [NOTEBOOKS]: ok({ value: [{ id: 'nb-a' }] }) })).defaultNotebookId).toBe('nb-a');
  });

  it('gives no working hours when the end time is absent', async () => {
    expect((await contextWith({ [MAILBOX]: ok({ workingHours: { startTime: '09:00' } }) })).tenantWorkingHours).toBeUndefined();
  });

  it('gives the working hours without a time zone when the mailbox has none', async () => {
    const context = await contextWith({ [MAILBOX]: ok({ workingHours: { startTime: '09:00', endTime: '17:00' } }) });
    expect(context.tenantWorkingHours).toEqual({ start: '09:00', end: '17:00', timeZone: undefined });
  });
});
