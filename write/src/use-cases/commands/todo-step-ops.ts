import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AppliedChange, GraphReads, ItemIds, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isGraphId, parseGraphId } from '../../domain/graph-id.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { TodoList, TodoTask } from './todo-reader.ts';
import { taskPath, unreadable } from './todo-reader.ts';

/*
 * The ops on the steps of a task (Graph calls them checklistItems): add,
 * change and delete. A step has no date of change, so the check that it is as
 * the plan read it compares its text and its check mark.
 */

/** The state of a step that a command sets and that its undo compares. */
export type StepState = { readonly displayName: string; readonly isChecked: boolean };

export type TodoStep = StepState & { readonly id: GraphId };

export const stepStateSchema = z.object({ displayName: z.string(), isChecked: z.boolean() });

const stepSchema = stepStateSchema.extend({ id: z.string() });

/** The Graph path of the steps of a task. */
export const stepsPath = (listId: string, taskId: string): string => `${taskPath(listId, taskId)}/checklistItems`;

const stepOf = (raw: unknown): Result<TodoStep, GraphError> => {
  const step = stepSchema.safeParse(raw);
  if (!step.success || !isGraphId(step.data.id)) return err(unreadable('a step of a To Do task'));
  return ok({ id: step.data.id, displayName: step.data.displayName, isChecked: step.data.isChecked });
};

/** The step `raw` of the task, after the check of the id. */
export const readStep = async (reads: GraphReads, task: TodoTask, raw: string): Promise<Result<TodoStep, GraphError>> => {
  const id = parseGraphId(raw, '--step-id');
  if (!id.ok) return id;
  const read = await reads.get(`${stepsPath(task.listId, task.id)}/${id.value}`);
  return read.ok ? stepOf(read.value) : read;
};

const stateOf = (step: StepState): StepState => ({ displayName: step.displayName, isChecked: step.isChecked });

const sameState = (a: StepState, b: StepState): boolean => a.displayName === b.displayName && a.isChecked === b.isChecked;

const stepBefore = (list: TodoList, task: TodoTask, step: TodoStep): ItemIds => ({
  listId: list.id,
  list: list.displayName,
  taskId: task.id,
  title: task.title,
  stepId: step.id,
  step: step.displayName,
  checked: String(step.isChecked),
});

/** Adds the step to the task, as the change `change`; the undo deletes it if it did not change. */
export const addStepWrite = async (graph: WriteGraph, list: TodoList, task: TodoTask, step: StepState, change?: PlannedChange): Promise<Result<AppliedChange, GraphError>> => {
  const target = stepsPath(list.id, task.id);
  const added = await graph.post(target, step);
  if (!added.ok) return added;
  const made = stepOf(added.value);
  if (!made.ok) return made;
  const ids = { listId: list.id, taskId: task.id, stepId: made.value.id };
  return ok({
    ...(change ?? { action: 'add-step', target, before: { listId: list.id, taskId: task.id } }),
    after: ids,
    undo: { kind: 'delete-created-todo-step', data: { ...ids, set: stateOf(made.value) } },
  });
};

/** Adds a step to the task. */
export const addStepOp = (list: TodoList, task: TodoTask, step: StepState): TodoOp => {
  const change: PlannedChange = {
    action: 'add-step',
    target: stepsPath(list.id, task.id),
    before: { listId: list.id, list: list.displayName, taskId: task.id, title: task.title },
  };
  return { change, write: async (graph) => outcomeOf(await addStepWrite(graph, list, task, step, change)) };
};

const recheckStep = async (reads: GraphReads, task: TodoTask, step: TodoStep): Promise<Result<boolean, GraphError>> => {
  const now = await readStep(reads, task, step.id);
  return now.ok ? ok(sameState(now.value, step)) : now;
};

/** Changes the text or the check mark of the step. */
export const updateStepOp = (list: TodoList, task: TodoTask, step: TodoStep, set: Partial<StepState>): TodoOp => {
  const change: PlannedChange = { action: 'update-step', target: `${stepsPath(list.id, task.id)}/${step.id}`, before: stepBefore(list, task, step) };
  return {
    change,
    recheck: async (reads) => recheckStep(reads, task, step),
    write: async (graph) => {
      const written = await graph.patch(change.target, set);
      if (!written.ok) return outcomeOf(written);
      const now = stepOf(written.value);
      if (!now.ok) return outcomeOf(now);
      const ids = { listId: list.id, taskId: task.id, stepId: step.id };
      return outcomeOf(
        ok({
          ...change,
          after: { ...ids, checked: String(now.value.isChecked) },
          undo: { kind: 'restore-todo-step', data: { ...ids, before: stateOf(step), set: stateOf(now.value) } },
        })
      );
    },
  };
};

/** Deletes the step. The undo adds it again, with a new id. */
export const deleteStepOp = (list: TodoList, task: TodoTask, step: TodoStep): TodoOp => {
  const change: PlannedChange = { action: 'delete-step', target: `${stepsPath(list.id, task.id)}/${step.id}`, before: stepBefore(list, task, step) };
  return {
    change,
    recheck: async (reads) => recheckStep(reads, task, step),
    write: async (graph) => {
      const deleted = await graph.delete(change.target);
      if (!deleted.ok) return outcomeOf(deleted);
      return outcomeOf(ok({ ...change, after: {}, undo: { kind: 'recreate-todo-step', data: { listId: list.id, taskId: task.id, step: stateOf(step) } } }));
    },
  };
};
