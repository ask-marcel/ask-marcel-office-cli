import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectOnlyOptions } from './odata-query.ts';

// Graph's `/me/planner/plans` honors `$select` server-side; every other
// OData passthrough is silently dropped (verified live — `--top 1` still
// returns the full list, `--filter` / `--orderby` are no-ops, combining
// multiple OData params trips an internal Graph fault). Expose only
// `--select` to keep the manifest honest.
const baseSchema = z.object({}).strict();
const { execute, schema } = buildPickODataListCommand(() => '/me/planner/plans', baseSchema, ['select']);

const meta: CommandMeta = {
  summary:
    'List every Microsoft Planner plan the signed-in user has access to (across every group). Use this to discover plan IDs without needing an existing task as the entry point. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — slice / sort client-side.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/plans',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-plans',
  options: [...selectOnlyOptions],
  example: 'ask-marcel list-planner-plans',
  responseShape: 'collection of Microsoft Graph `plannerPlan` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
