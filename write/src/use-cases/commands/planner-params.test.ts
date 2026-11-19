import { describe, expect, it } from 'bun:test';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { ANN, createPlannerStoreFake, ME } from '../../test-helpers/planner-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * The params of the Planner commands: a missing or blank value, a value
 * outside its list and an id that is not a Graph id are refused before any
 * request, with a refusal of the `validation` source.
 */

const files = createLocalBytesFake({ 'a.txt': utf8Bytes('x') });

type Case = readonly [string, Record<string, string>];

const errorOf = async (name: string, params: Record<string, string>): Promise<{ readonly planner: ReturnType<typeof createPlannerStoreFake>; readonly error: unknown }> => {
  const planner = createPlannerStoreFake();
  const run = await runWrite(planner.graph, name, params, { files });
  return { planner, error: run['error'] };
};

const MISSING: ReadonlyArray<Case> = [
  ['create-planner-bucket', { plan: 'plan-1' }],
  ['update-planner-bucket', { name: 'x' }],
  ['delete-planner-bucket', {}],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a' }],
  ['update-planner-task', { title: 'x' }],
  ['delete-planner-task', {}],
  ['update-planner-task-details', { description: 'x' }],
  ['add-planner-checklist-item', { taskId: 'task-1' }],
  ['update-planner-checklist-item', { taskId: 'task-1', checked: 'true' }],
  ['remove-planner-checklist-item', { taskId: 'task-1' }],
  ['add-planner-task-file', { taskId: 'task-1' }],
  ['add-planner-task-link', { taskId: 'task-1' }],
  ['remove-planner-task-link', { taskId: 'task-1' }],
  ['update-planner-plan', { name: 'x' }],
];

const BLANK: ReadonlyArray<Case> = [
  ['create-planner-bucket', { plan: '  ', name: 'x' }],
  ['create-planner-bucket', { plan: 'plan-1', name: '  ' }],
  ['update-planner-bucket', { bucketId: 'bucket-a', name: '  ' }],
  ['update-planner-bucket', { bucketId: 'bucket-a', after: '  ' }],
  ['create-planner-task', { plan: '  ', bucket: 'bucket-a', title: 'x' }],
  ['create-planner-task', { plan: 'plan-1', bucket: '  ', title: 'x' }],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: '  ' }],
  ['update-planner-task', { taskId: 'task-1', title: '  ' }],
  ['update-planner-task', { taskId: 'task-1', bucket: '  ' }],
  ['add-planner-checklist-item', { taskId: 'task-1', title: '  ' }],
  ['update-planner-checklist-item', { taskId: 'task-1', item: 'c1', title: '  ' }],
  ['add-planner-task-file', { taskId: 'task-1', file: 'a.txt', alias: '  ' }],
  ['add-planner-task-link', { taskId: 'task-1', url: 'https://example.com', alias: '  ' }],
  ['update-planner-plan', { plan: '  ', name: 'x' }],
  ['update-planner-plan', { plan: 'plan-1', name: '  ' }],
];

const OUT_OF_LIST: ReadonlyArray<Case> = [
  ['update-planner-bucket', { bucketId: 'bucket-a', first: 'yes' }],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', notifyOthers: 'yes' }],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', status: 'done' }],
  ['update-planner-task', { taskId: 'task-1', notifyOthers: 'yes', title: 'x' }],
  ['delete-planner-task', { taskId: 'task-1', notifyOthers: 'yes' }],
  ['update-planner-task-details', { taskId: 'task-1', preview: 'none' }],
  ['add-planner-checklist-item', { taskId: 'task-1', title: 'x', checked: 'yes' }],
];

const BAD_ID: ReadonlyArray<Case> = [
  ['update-planner-bucket', { bucketId: 'a/b', name: 'x' }],
  ['update-planner-task', { taskId: 'a/b', title: 'x' }],
  ['delete-planner-task', { taskId: 'a/b' }],
  ['update-planner-task-details', { taskId: 'a/b', description: 'x' }],
  ['add-planner-checklist-item', { taskId: 'a/b', title: 'x' }],
  ['update-planner-checklist-item', { taskId: 'a/b', item: 'c1', checked: 'true' }],
  ['remove-planner-checklist-item', { taskId: 'a/b', item: 'c1' }],
  ['add-planner-task-file', { taskId: 'a/b', file: 'a.txt' }],
  ['add-planner-task-link', { taskId: 'a/b', url: 'https://example.com' }],
  ['remove-planner-task-link', { taskId: 'a/b', url: 'https://example.com' }],
];

describe('the params of the Planner commands', () => {
  it.each([...MISSING, ...BLANK, ...OUT_OF_LIST])('refuses %s with %o before any request', async (name, params) => {
    const { planner, error } = await errorOf(name, params);
    expect(error).toMatchObject({ source: 'validation' });
    expect(planner.calls).toEqual([]);
  });

  it.each(BAD_ID)('refuses %s with an id that is not a Graph id, %o, before any request', async (name, params) => {
    const { planner, error } = await errorOf(name, params);
    expect(error).toMatchObject({ code: 'invalid_id', source: 'validation', message: expect.stringContaining('"a/b" is not a Graph id') });
    expect(planner.calls).toEqual([]);
  });

  it('takes false for each true-or-false option', async () => {
    const planner = createPlannerStoreFake();
    const dry = { dryRun: true };
    expect(await runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', assign: 'me', notifyOthers: 'false' }, dry)).toMatchObject({
      dryRun: true,
    });
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', title: 'x', notifyOthers: 'false' }, dry)).toMatchObject({ dryRun: true });
    expect(await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1', notifyOthers: 'false' }, dry)).toMatchObject({ dryRun: true });
    expect(await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'x', checked: 'false' }, dry)).toMatchObject({ dryRun: true });
  });

  it('gives the words of the refusal of a field', async () => {
    const message = async (params: Record<string, string>): Promise<unknown> =>
      ((await errorOf('update-planner-task', { taskId: 'task-1', ...params })).error as { message: string }).message;
    expect(await message({ due: '2026-02-30' })).toContain('must be a day of the calendar in the form YYYY-MM-DD, or none to remove it');
    expect(await message({ percent: '101' })).toContain('must be a whole number from 0 to 100');
    expect(await message({ priority: 'high' })).toContain('must be urgent, important, medium, low, or a whole number from 0 to 10');
    for (const percent of ['1100', '1001', '-1', '5.5']) expect(await message({ percent })).toContain('must be a whole number from 0 to 100');
    expect(await message({ status: 'completed', percent: '10' })).toBe('--status and --percent both set the progress of the task: give one of them.');
  });
});

describe('the people of a Planner task', () => {
  const assign = async (spec: string, extra: Record<string, string> = {}): Promise<Record<string, unknown>> => {
    const planner = createPlannerStoreFake();
    return runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', assign: spec, ...extra }, { dryRun: true });
  };

  it('refuses a person that is not "me", a mail address or a user id, and names the item', async () => {
    for (const spec of [
      'me,a@b',
      'me,a@b.c@d.e',
      'me,a b@c.d',
      'me,a@c.d e',
      'me,x11111111-1111-4111-8111-111111111111',
      'me,11111111-1111-4111-8111-111111111111x',
      'me,a/b@c.d',
    ]) {
      expect(await assign(spec)).toMatchObject({
        error: { code: 'planner_person_invalid', source: 'validation', message: '--assign: item 2 is not "me", a mail address or a user id.' },
      });
    }
  });

  it('refuses a person that the tenant does not have, and gives the error of another failed read', async () => {
    expect(await assign('zed@example.com')).toMatchObject({
      error: { code: 'planner_person_not_found', message: '--assign: item 1: no user of the tenant has this mail address or id.' },
    });
    const planner = createPlannerStoreFake();
    planner.failures.set('GET /users/ann%40example.com', { type: 'api_error', status: 503, message: 'Try again later.' });
    expect(await runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', assign: 'ann@example.com', notifyOthers: 'true' })).toMatchObject(
      {
        error: { message: 'Try again later.' },
      }
    );
  });

  it('takes the user by its id or its mail with no consent, once, and another person by its user id with consent', async () => {
    expect(await assign(`${ME.id},ME,me`)).toMatchObject({ changes: [{ before: { assign: 'Robin Doe (you)' } }] });
    expect(await assign(ANN.id, { notifyOthers: 'true' })).toMatchObject({ changes: [{ before: { assign: 'Ann Lee' } }] });
  });

  it('shows the id of a person with no display name', async () => {
    const planner = createPlannerStoreFake();
    planner.answers.set('GET /users/ann%40example.com', { id: ANN.id, displayName: null });
    const dry = await runWrite(
      planner.graph,
      'create-planner-task',
      { plan: 'plan-1', bucket: 'bucket-a', title: 'x', assign: 'ann@example.com', notifyOthers: 'true' },
      { dryRun: true }
    );
    expect(dry).toMatchObject({ changes: [{ before: { assign: ANN.id } }] });
  });

  it('shows the people added and removed, leaves out a person assigned already, and reads no user for a change of no person', async () => {
    const planner = createPlannerStoreFake();
    const dry = await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', assign: 'ann@example.com', unassign: 'me', notifyOthers: 'true' }, { dryRun: true });
    expect(dry).toMatchObject({ changes: [{ before: { assign: 'Ann Lee', unassign: 'Robin Doe (you)' } }] });
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', assign: 'me' })).toMatchObject({ error: { code: 'planner_nothing_to_change' } });
    const quiet = createPlannerStoreFake();
    await runWrite(quiet.graph, 'update-planner-task', { taskId: 'task-1', title: 'New' }, { dryRun: true });
    expect(quiet.calls.filter((call) => call.startsWith('GET /me') || call.startsWith('GET /users'))).toEqual([]);
  });

  it('gives the words of the refusals of the people', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', unassign: 'x y' })).toMatchObject({
      error: { message: '--unassign: item 1 is not "me", a mail address or a user id.' },
    });
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', assign: 'me', unassign: 'me' })).toMatchObject({
      error: { code: 'planner_assign_and_unassign', source: 'validation', message: '--assign and --unassign name the same person: give the person in one of them.' },
    });
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1' })).toMatchObject({
      error: { message: 'The params change nothing of the task: give a field, --labels, --bucket, --assign or --unassign.' },
    });
  });
});
