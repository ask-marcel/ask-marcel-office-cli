import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';

/**
 * Microsoft To Do behind a fake write graph, for the tests of the To Do
 * commands. It answers the requests these commands make, under the lists
 * path of To Do: the lists, the tasks of a list, and the steps
 * (`checklistItems`), the linked resources and the attachments of a task; and
 * `GET /me/mailboxSettings/timeZone`. A new item gets the id `<kind>-<n>`. A
 * change of a task, or of an item under it, gives the task a new
 * `lastModifiedDateTime`, as To Do does.
 *
 * `calls` logs each request as `<METHOD> <path>` (an upload as `UPLOAD <path of
 * the attachments>`). `failures` makes a request fail: the key is `<METHOD>
 * <path>` (the path without its query), the value the error. `edit` changes a
 * task as the To Do app would.
 */
export type Item = Record<string, unknown> & { readonly id: string };

export type TodoStoreFake = {
  readonly graph: WriteGraph;
  readonly lists: Map<string, Item>;
  /** Each task holds its `listId`; each step, link and file holds its `taskId`. */
  readonly tasks: Map<string, Item>;
  readonly steps: Map<string, Item>;
  readonly links: Map<string, Item>;
  readonly files: Map<string, Item>;
  readonly calls: string[];
  readonly failures: Map<string, GraphError>;
  readonly edit: (taskId: string, fields: Record<string, unknown>) => void;
};

export const TIME_ZONE = 'Singapore Standard Time';

const notFound: GraphError = { type: 'api_error', status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.' };

const CHILDREN = { checklistItems: 'steps', linkedResources: 'links', attachments: 'files' } as const;

type ChildKind = (typeof CHILDREN)[keyof typeof CHILDREN];

type Store = Pick<TodoStoreFake, 'lists' | 'tasks' | 'steps' | 'links' | 'files'>;

type Path = { readonly list?: string; readonly task?: string; readonly child?: ChildKind; readonly item?: string };

const parse = (path: string): Path => {
  const [, , , , list, , task, child, item] = (path.split('?')[0] ?? '').split('/');
  const kind = child !== undefined && Object.hasOwn(CHILDREN, child) ? CHILDREN[child as keyof typeof CHILDREN] : undefined;
  return { list, task, child: kind, item };
};

const found = (item: Item | undefined): Result<unknown, GraphError> => (item === undefined ? err(notFound) : ok(structuredClone(item)));

// Without its bytes, as Graph lists the attachments of a task.
const listed = (kind: ChildKind, item: Item): Item => {
  if (kind !== 'files') return item;
  return Object.fromEntries(Object.entries(item).filter(([key]) => key !== 'contentBytes')) as Item;
};

export const createTodoStoreFake = (
  seed: { readonly lists?: ReadonlyArray<Item>; readonly tasks?: ReadonlyArray<Item>; readonly steps?: ReadonlyArray<Item> } = {}
): TodoStoreFake => {
  const store: Store = {
    lists: new Map((seed.lists ?? []).map((list) => [list.id, list])),
    tasks: new Map((seed.tasks ?? []).map((task) => [task.id, task])),
    steps: new Map((seed.steps ?? []).map((step) => [step.id, step])),
    links: new Map(),
    files: new Map(),
  };
  const calls: string[] = [];
  const failures = new Map<string, GraphError>();
  let clock = 0;
  const newId = (kind: string): string => `${kind}-${++clock}`;
  const stamp = (): string => new Date(Date.UTC(2026, 10, 18, 10, 0, ++clock)).toISOString();
  const touch = (taskId: string): void => {
    const task = store.tasks.get(taskId);
    if (task !== undefined) store.tasks.set(taskId, { ...task, lastModifiedDateTime: stamp() });
  };
  const taskOf = (path: Path): Item | undefined => {
    const task = store.tasks.get(path.task ?? '');
    return task?.['listId'] === path.list ? task : undefined;
  };
  const childOf = (path: Path): Item | undefined => {
    const item = path.child === undefined ? undefined : store[path.child].get(path.item ?? '');
    return item?.['taskId'] === path.task ? item : undefined;
  };
  const failure = (method: string, path: string): GraphError | undefined => failures.get(`${method} ${path.split('?')[0] ?? ''}`);

  const getChildren = (path: Path, kind: ChildKind): Result<unknown, GraphError> =>
    taskOf(path) === undefined ? err(notFound) : ok({ value: [...store[kind].values()].filter((item) => item['taskId'] === path.task).map((item) => listed(kind, item)) });

  const getAt = (raw: string): Result<unknown, GraphError> => {
    if (raw.startsWith('/me/mailboxSettings/timeZone')) return ok({ value: TIME_ZONE });
    const path = parse(raw);
    if (path.list === undefined) return ok({ value: [...store.lists.values()] });
    if (path.task === undefined) return found(store.lists.get(path.list));
    if (path.child === undefined) return found(taskOf(path));
    return path.item === undefined ? getChildren(path, path.child) : found(childOf(path));
  };

  const addTask = (path: Path, body: Record<string, unknown>): Result<unknown, GraphError> => {
    if (path.list === undefined || !store.lists.has(path.list)) return err(notFound);
    const task: Item = { status: 'notStarted', importance: 'normal', ...body, id: newId('task'), listId: path.list, lastModifiedDateTime: stamp() };
    store.tasks.set(task.id, task);
    return ok(structuredClone(task));
  };

  const addChild = (path: Path, kind: ChildKind, body: Record<string, unknown>): Result<unknown, GraphError> => {
    if (taskOf(path) === undefined) return err(notFound);
    const item: Item = { ...body, id: newId(kind), taskId: path.task };
    store[kind].set(item.id, item);
    touch(path.task ?? '');
    return ok(structuredClone(item));
  };

  const postAt = (raw: string, body: Record<string, unknown>): Result<unknown, GraphError> => {
    const path = parse(raw);
    if (path.list === undefined) {
      const list: Item = { ...body, id: newId('list') };
      store.lists.set(list.id, list);
      return ok(structuredClone(list));
    }
    return path.child === undefined ? addTask(path, body) : addChild(path, path.child, body);
  };

  const patchAt = (raw: string, body: Record<string, unknown>): Result<unknown, GraphError> => {
    const path = parse(raw);
    if (path.task === undefined) {
      const list = store.lists.get(path.list ?? '');
      if (list === undefined) return err(notFound);
      store.lists.set(list.id, { ...list, ...body });
      return found(store.lists.get(list.id));
    }
    const item = path.child === undefined ? taskOf(path) : childOf(path);
    if (item === undefined) return err(notFound);
    const map = path.child === undefined ? store.tasks : store[path.child];
    map.set(item.id, { ...item, ...body });
    touch(path.task);
    return found(map.get(item.id));
  };

  const deleteAt = (raw: string): Result<unknown, GraphError> => {
    const path = parse(raw);
    const item = path.child === undefined ? taskOf(path) : childOf(path);
    if (item === undefined) return err(notFound);
    if (path.child === undefined) store.tasks.delete(item.id);
    else {
      store[path.child].delete(item.id);
      touch(path.task ?? '');
    }
    return ok({});
  };

  const createLogged =
    <A extends unknown[]>(method: string, answer: (path: string, ...rest: A) => Result<unknown, GraphError>) =>
    async (path: string, ...rest: A): Promise<Result<unknown, GraphError>> => {
      calls.push(`${method} ${path}`);
      const failed = failure(method, path);
      return failed === undefined ? answer(path, ...rest) : err(failed);
    };

  const graph = fakeWriteGraph({
    get: createLogged('GET', (path) => getAt(path)),
    post: createLogged('POST', (path, body: unknown) => postAt(path, body as Record<string, unknown>)),
    patch: createLogged('PATCH', (path, body: unknown) => patchAt(path, body as Record<string, unknown>)),
    delete: createLogged('DELETE', (path) => deleteAt(path)),
    uploadAttachment: async (target, file) => {
      if (target.kind !== 'todo-task') return err(notFound);
      const at = `/me/todo/lists/${target.listId}/tasks/${target.taskId}/attachments`;
      calls.push(`UPLOAD ${at}`);
      const failed = failures.get(`UPLOAD ${at}`);
      if (failed !== undefined) return err(failed);
      const bytes = Buffer.from(file.bytes).toString('base64');
      const added = addChild(parse(at), 'files', {
        name: file.name,
        contentType: file.contentType ?? 'application/octet-stream',
        size: file.bytes.byteLength,
        contentBytes: bytes,
      });
      if (!added.ok) return added;
      return ok({ id: (added.value as Item).id, name: file.name, size: file.bytes.byteLength });
    },
  });

  return {
    graph,
    ...store,
    calls,
    failures,
    edit: (taskId, fields) => {
      const task = store.tasks.get(taskId);
      if (task !== undefined) store.tasks.set(taskId, { ...task, ...fields, lastModifiedDateTime: stamp() });
    },
  };
};

/** Two lists (`Tasks`, `Later`), one task in `Tasks` with a due day, a note and a category, and one step of it. */
export const TODO_SEED = {
  lists: [
    { id: 'list-a', displayName: 'Tasks' },
    { id: 'list-b', displayName: 'Later' },
  ],
  tasks: [
    {
      id: 'task-a',
      listId: 'list-a',
      title: 'Send the report',
      status: 'notStarted',
      importance: 'normal',
      categories: ['Blue'],
      body: { content: 'The note', contentType: 'text' },
      dueDateTime: { dateTime: '2026-11-19T16:00:00.0000000', timeZone: 'UTC' },
      isReminderOn: false,
      createdDateTime: '2026-11-01T08:00:00Z',
      lastModifiedDateTime: '2026-11-17T08:00:00Z',
    },
  ],
  steps: [{ id: 'step-a', taskId: 'task-a', displayName: 'Read the draft', isChecked: false }],
} as const satisfies Parameters<typeof createTodoStoreFake>[0];

/** A To Do store with the seed of this file. */
export const createSeededTodoStore = (): TodoStoreFake => createTodoStoreFake(structuredClone(TODO_SEED));
