import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../domain/result.ts';
import { writeCommandOf } from '../use-cases/commands/command-graph.ts';
import type { PlanWrite } from '../use-cases/commands/write-plan.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeWriteGraph } from '../test-helpers/write-graph-fake.ts';
import { runRegistryCommand } from './run-registry-command.ts';

/*
 * A plan of a batch passes over the items that need no change or that are no
 * longer there, and says why. The dry-run shows them beside the changes, so the
 * user sees each item of the batch.
 */

const createPlan =
  (skipped: boolean): PlanWrite =>
  async () =>
    ok({
      changes: [{ action: 'tag-widget', target: '/widgets/w1', before: { widgetId: 'w1' } }],
      ...(skipped ? { skipped: [{ target: '/widgets/w2', before: { widgetId: 'w2' }, reason: 'The widget has the tag already.' }] } : {}),
      apply: async () => ok({ value: {}, changes: [] }),
    });

const tagWidgets = (skipped: boolean): ReturnType<typeof writeCommandOf> =>
  writeCommandOf({
    schema: z.object({}),
    execute: async () => ok({}),
    meta: { ...fakeCommand().meta, effect: 'mail-state', category: 'mail', graphMethod: 'PATCH', undo: 'removes the tag' },
    plan: createPlan(skipped),
  });

const dryRun = async (skipped: boolean): Promise<unknown> => {
  const command = tagWidgets(skipped);
  const deps = { graph: fakeWriteGraph(), fs: createFileSystemFake(), commandPackage: fakeCommandPackage({ registry: { 'tag-widgets': command } }) };
  const result = await runRegistryCommand(deps, { name: 'tag-widgets', command, params: {}, surface: 'cli', dryRun: true });
  return result.ok ? result.value : result.error;
};

describe('the dry-run of a batch', () => {
  it('shows the items that the plan passes over, with the reason', async () => {
    expect(await dryRun(true)).toEqual({
      dryRun: true,
      command: 'tag-widgets',
      changes: [{ action: 'tag-widget', target: '/widgets/w1', before: { widgetId: 'w1' } }],
      skipped: [{ target: '/widgets/w2', before: { widgetId: 'w2' }, reason: 'The widget has the tag already.' }],
    });
  });

  it('has no skipped list when the plan passes over nothing', async () => {
    expect(await dryRun(false)).toEqual({ dryRun: true, command: 'tag-widgets', changes: [{ action: 'tag-widget', target: '/widgets/w1', before: { widgetId: 'w1' } }] });
  });
});
