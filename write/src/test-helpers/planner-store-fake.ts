import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import type { PatchOptions, WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';

/**
 * Planner behind a fake write graph, for the tests of the Planner commands.
 * As Planner does: each item has an ETag that each write changes; a PATCH or
 * a DELETE with no `If-Match` is refused (400) and one with an old ETag too
 * (412); a PATCH answers the item only with `representation`; the `notes` of
 * the details go to beta only, and give the plain `description`; a change of
 * the details changes the counters of the task; a delete of a bucket deletes
 * its tasks; a new task gets empty details. A map field of a PATCH
 * (assignments, labels, checklist, references, label names) changes only the
 * entries it names, and `null` removes one. `/me`, `/users/{mail or id}`,
 * the plans of the user and the OneDrive folder of the attachments are there.
 *
 * `calls` logs `<METHOD> <path>` (with ` [beta]` for a beta PATCH), `failures`
 * and `answers` change the answer of `<METHOD> <path>`, `afterRead` runs after
 * each GET with its count, and `edit` changes an item as a person would.
 */
export type Item = Record<string, unknown>;

export type PlannerStoreFake = {
  readonly graph: WriteGraph;
  readonly items: Map<string, Item>;
  readonly calls: string[];
  readonly failures: Map<string, GraphError>;
  readonly answers: Map<string, unknown>;
  readonly hooks: { afterRead?: (count: number, path: string) => void };
  readonly edit: (path: string, fields: Item) => void;
};

export const ME = { id: '11111111-1111-4111-8111-111111111111', displayName: 'Robin Doe' };
export const ANN = { id: '22222222-2222-4222-8222-222222222222', displayName: 'Ann Lee', mail: 'ann@example.com' };

const notFound: GraphError = { type: 'api_error', status: 404, code: 'NotFound', message: 'The requested item is not found.' };
const precondition: GraphError = { type: 'api_error', status: 412, code: 'PreconditionFailed', message: 'The If-Match header does not match.' };
const badRequest = (message: string): GraphError => ({ type: 'api_error', status: 400, code: 'BadRequest', message });

const MAPS = ['assignments', 'appliedCategories', 'checklist', 'references', 'categoryDescriptions'];

const merged = (item: Item, body: Item): Item => {
  const next: Item = { ...item };
  for (const [key, value] of Object.entries(body)) {
    if (!MAPS.includes(key)) next[key] = value;
    else next[key] = Object.fromEntries(Object.entries({ ...(item[key] as Item | undefined), ...(value as Item) }).filter(([, entry]) => entry !== null));
  }
  return next;
};

const textOf = (html: string): string =>
  html
    .split('<')
    .map((part, index) => (index === 0 ? part : part.slice(part.indexOf('>') + 1)))
    .join('');

/** The seed: a plan of the user with two labels and two buckets, a task with details, and a plan of a group. */
export const seedPlanner = (): Array<[string, Item]> => [
  ['/planner/plans/plan-1', { id: 'plan-1', title: 'Launch', container: { type: 'user' } }],
  ['/planner/plans/plan-1/details', { id: 'plan-1', categoryDescriptions: { category1: 'Red', category2: 'Blue' } }],
  ['/planner/plans/plan-g', { id: 'plan-g', title: 'Team', container: { type: 'group' } }],
  ['/planner/buckets/bucket-a', { id: 'bucket-a', name: 'To do', planId: 'plan-1', orderHint: 'a' }],
  ['/planner/buckets/bucket-b', { id: 'bucket-b', name: 'Doing', planId: 'plan-1', orderHint: 'b' }],
  [
    '/planner/tasks/task-1',
    {
      id: 'task-1',
      planId: 'plan-1',
      bucketId: 'bucket-a',
      title: 'Write the report',
      percentComplete: 0,
      priority: 5,
      dueDateTime: '2026-11-20T10:00:00Z',
      assignments: { [ME.id]: { orderHint: 'x' } },
      appliedCategories: { category1: true },
      checklistItemCount: 2,
      referenceCount: 1,
      hasDescription: true,
    },
  ],
  [
    '/planner/tasks/task-1/details',
    {
      id: 'task-1',
      description: 'The text',
      previewType: 'automatic',
      checklist: { c1: { title: 'Draft', isChecked: true, orderHint: '2' }, c2: { title: 'Send', isChecked: false, orderHint: '1' } },
      references: { 'https%3A//example%2Ecom/spec': { alias: 'Spec', type: 'Other' } },
    },
  ],
];

export const createPlannerStoreFake = (seed: Array<[string, Item]> = seedPlanner()): PlannerStoreFake => {
  let version = 0;
  const etag = (): string => `W/"v${++version}"`;
  const items = new Map<string, Item>(seed.map(([path, item]) => [path, { ...item, '@odata.etag': etag() }]));
  const calls: string[] = [];
  const failures = new Map<string, GraphError>();
  const answers = new Map<string, unknown>();
  const hooks: PlannerStoreFake['hooks'] = {};
  let made = 0;
  let reads = 0;

  const put = (path: string, item: Item): Item => {
    const kept = { ...item, '@odata.etag': etag() };
    items.set(path, kept);
    return kept;
  };

  const counters = (taskPath: string, details: Item): void => {
    const task = items.get(taskPath);
    const count = (key: string): number => Object.keys((details[key] as Item | undefined) ?? {}).length;
    if (task !== undefined) put(taskPath, { ...task, checklistItemCount: count('checklist'), referenceCount: count('references'), hasDescription: details['description'] !== '' });
  };

  const list = (path: string): Result<unknown, GraphError> | undefined => {
    const values = (prefix: string, match: (item: Item) => boolean): Result<unknown, GraphError> =>
      ok({ value: [...items].filter(([key, item]) => key.startsWith(prefix) && key.split('/').length === 4 && match(item)).map(([, item]) => item) });
    if (path === '/me/planner/plans') return values('/planner/plans/', () => true);
    const [, , kind, id, sub] = path.split('/');
    if (kind === 'plans' && sub === 'buckets') return values('/planner/buckets/', (item) => item['planId'] === id);
    if (kind === 'buckets' && sub === 'tasks') return values('/planner/tasks/', (item) => item['bucketId'] === id);
    return undefined;
  };

  const user = (path: string): Result<unknown, GraphError> | undefined => {
    if (path === '/me') return ok(ME);
    if (!path.startsWith('/users/')) return undefined;
    const spec = decodeURIComponent(path.slice('/users/'.length));
    return spec === ANN.id || spec === ANN.mail ? ok({ id: ANN.id, displayName: ANN.displayName }) : err(notFound);
  };

  const getAt = (path: string): Result<unknown, GraphError> => {
    const item = items.get(path);
    return list(path) ?? user(path) ?? (item === undefined ? err(notFound) : ok(structuredClone(item)));
  };

  const create = (path: string, body: Item): Result<unknown, GraphError> => {
    const kind = path === '/planner/tasks' ? 'tasks' : 'buckets';
    const id = `${kind === 'tasks' ? 'task' : 'bucket'}-new-${++made}`;
    const item = put(`/planner/${kind}/${id}`, { ...body, id, ...(kind === 'buckets' ? { orderHint: String(body['orderHint'] ?? `new-${made}`) } : {}) });
    if (kind === 'tasks') put(`/planner/tasks/${id}/details`, { id, description: '', previewType: 'automatic', checklist: {}, references: {} });
    return ok(structuredClone(item));
  };

  const checkMatch = (path: string, options: PatchOptions | undefined): Result<Item, GraphError> => {
    const item = items.get(path);
    if (item === undefined) return err(notFound);
    if (options?.ifMatch === undefined) return err(badRequest('The If-Match header is missing.'));
    return options.ifMatch === item['@odata.etag'] ? ok(item) : err(precondition);
  };

  const detailsBody = (path: string, body: Item, beta: boolean): Result<Item, GraphError> => {
    if (!path.endsWith('/details') || !Object.hasOwn(body, 'notes')) return ok(body);
    if (!beta) return err(badRequest('The property notes does not exist on type plannerTaskDetails.'));
    if (Object.hasOwn(body, 'description')) return err(badRequest('Setting both description and notes is an error.'));
    const content = String((body['notes'] as Item)['content'] ?? '');
    return ok({ ...body, notes: { contentType: 'html', content }, description: textOf(content) });
  };

  const change = (path: string, body: Item, options: PatchOptions | undefined): Result<unknown, GraphError> => {
    const found = checkMatch(path, options);
    if (!found.ok) return found;
    const written = detailsBody(path, body, options?.api === 'beta');
    if (!written.ok) return written;
    const kept = put(path, merged(found.value, written.value));
    if (path.startsWith('/planner/tasks/') && path.endsWith('/details')) counters(path.slice(0, -'/details'.length), kept);
    return ok(options?.representation === true ? structuredClone(items.get(path)) : undefined);
  };

  const remove = (path: string, options: PatchOptions | undefined): Result<unknown, GraphError> => {
    const found = checkMatch(path, options);
    if (!found.ok) return found;
    for (const key of [...items.keys()]) if (key === path || key.startsWith(`${path}/`)) items.delete(key);
    if (path.startsWith('/planner/buckets/')) for (const [key, item] of [...items]) if (item['bucketId'] === path.split('/')[3]) items.delete(key);
    return ok(undefined);
  };

  const answer = (call: string, run: () => Result<unknown, GraphError>): Result<unknown, GraphError> => {
    calls.push(call);
    const key = call.replace(' [beta]', '');
    const failure = failures.get(key);
    if (failure !== undefined) return err(failure);
    const result = run();
    return answers.has(key) ? ok(answers.get(key)) : result;
  };

  const graph = fakeWriteGraph({
    get: async (raw) => {
      const path = raw.split('?')[0] ?? '';
      const result = answer(`GET ${path}`, () => getAt(path));
      hooks.afterRead?.(++reads, path);
      return result;
    },
    post: async (path, body) => answer(`POST ${path}`, () => create(path, body as Item)),
    patch: async (path, body, options) => answer(`PATCH ${path}${options?.api === 'beta' ? ' [beta]' : ''}`, () => change(path, body as Item, options)),
    delete: async (path, options) => answer(`DELETE ${path}`, () => remove(path, options)),
  });

  const edit = (path: string, fields: Item): void => {
    const item = items.get(path);
    if (item !== undefined) put(path, merged(item, fields));
  };

  return { graph, items, calls, failures, answers, hooks, edit };
};
