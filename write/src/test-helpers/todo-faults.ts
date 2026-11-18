import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { TodoStoreFake } from './todo-store-fake.ts';
import { createSeededTodoStore } from './todo-store-fake.ts';
import { runWrite, undoWrite } from './run-write.ts';

/**
 * The fault sweeps of the To Do commands: a scenario runs once to count its
 * Graph calls, then once for each call, with that call failed (`fail`) or
 * answered with a shape that the CLI cannot read (`garbage`). A command that
 * goes on after a failed read, or that names the wrong item in a refusal,
 * shows it in the outcome.
 */
export type Fault = 'fail' | 'garbage';

export type Scenario = {
  readonly command: string;
  readonly params: Record<string, string>;
  readonly prepare?: (todo: TodoStoreFake) => Promise<void>;
  readonly files?: LocalFileReader;
};

/** The error of the failed call `at`. */
export const failureAt = (at: number): string => `Graph failed at call ${at}.`;

type Faulty = { readonly graph: WriteGraph; readonly calls: () => number };

// The graph with the call `at` (from 0) answered by the fault; a negative `at` changes no call.
export const createFaultyGraph = (graph: WriteGraph, at: number, fault: Fault): Faulty => {
  let count = 0;
  const hit = (): boolean => count++ === at;
  const answer = (): Result<unknown, GraphError> => (fault === 'fail' ? err({ type: 'api_error', status: 500, message: failureAt(at) }) : ok({ bogus: true }));
  return {
    graph: {
      ...graph,
      get: async (path, headers) => (hit() ? answer() : graph.get(path, headers)),
      post: async (path, body) => (hit() ? answer() : graph.post(path, body)),
      patch: async (path, body, options) => (hit() ? answer() : graph.patch(path, body, options)),
      delete: async (path, options) => (hit() ? answer() : graph.delete(path, options)),
      uploadAttachment: async (target, file) => {
        if (!hit()) return graph.uploadAttachment(target, file);
        return fault === 'fail' ? err({ type: 'api_error', status: 500, message: failureAt(at) }) : ok({ name: file.name, size: 1 });
      },
    },
    calls: () => count,
  };
};

const prepared = async (scenario: Scenario): Promise<TodoStoreFake> => {
  const todo = createSeededTodoStore();
  await scenario.prepare?.(todo);
  return todo;
};

const optionsOf = (scenario: Scenario): { readonly files?: LocalFileReader } => (scenario.files === undefined ? {} : { files: scenario.files });

/** The outcome of the run of the scenario with the fault on the call `at` (a negative `at`: no fault), as JSON, and its count of calls. */
export const runWithFault = async (scenario: Scenario, at: number, fault: Fault): Promise<{ readonly text: string; readonly calls: number }> => {
  const faulty = createFaultyGraph((await prepared(scenario)).graph, at, fault);
  const result = await runWrite(faulty.graph, scenario.command, scenario.params, optionsOf(scenario));
  return { text: JSON.stringify(result), calls: faulty.calls() };
};

/** The outcome of the undo (applied) of a clean run of the scenario, with the fault on the call `at` of the undo. */
export const undoWithFault = async (scenario: Scenario, at: number, fault: Fault): Promise<{ readonly text: string; readonly calls: number }> => {
  const todo = await prepared(scenario);
  const journal = createRunJournalFake();
  await runWrite(todo.graph, scenario.command, scenario.params, { ...optionsOf(scenario), journal });
  const faulty = createFaultyGraph(todo.graph, at, fault);
  const report = await undoWrite(faulty.graph, journal, true);
  return { text: JSON.stringify(report), calls: faulty.calls() };
};

const UNREADABLE = /Graph gave ([^"]*?) in a shape that this CLI cannot read\./;

/** The thing that each call of the sweep names as unreadable when it gets a garbage answer, or `-` when the outcome names none. */
export const garbageSweep = async (run: (at: number, fault: Fault) => Promise<{ readonly text: string; readonly calls: number }>): Promise<ReadonlyArray<string>> => {
  const total = (await run(-1, 'garbage')).calls;
  const named: string[] = [];
  for (let at = 0; at < total; at += 1) named.push(UNREADABLE.exec((await run(at, 'garbage')).text)?.[1] ?? '-');
  return named;
};

/** The calls of the sweep whose failure the outcome does not show. */
export const unseenFailures = async (run: (at: number, fault: Fault) => Promise<{ readonly text: string; readonly calls: number }>): Promise<ReadonlyArray<number>> => {
  const total = (await run(-1, 'fail')).calls;
  const unseen: number[] = [];
  for (let at = 0; at < total; at += 1) if (!(await run(at, 'fail')).text.includes(failureAt(at))) unseen.push(at);
  return total === 0 ? [-1] : unseen;
};
