import type { GraphError } from '../../domain/graph-error.ts';
import type { Result } from '../../domain/result.ts';
import type { WriteGraph } from '../ports/graph.ts';
import type { LocalFileReader } from './read-local-file.ts';

/*
 * The contract of a write that can run as a dry-run and be undone (step 2 of
 * the write plan). A command declares its changes and how to reverse each one;
 * the dry-run, the run journal and the undo are generic (record-run.ts,
 * undo-run.ts). docs/WRITE-RUNS.md tells a command author how to use it.
 */

/** The ids of an item, for example `{ messageId: 'AAMk...', changeKey: 'CQAA...' }`. Strings only, so a record is plain JSON. */
export type ItemIds = Readonly<Record<string, string>>;

/** One change a command will make, as a dry-run shows it. */
export type PlannedChange = {
  /** What the change does, in the command's words, for example `create-draft`. */
  readonly action: string;
  /** The Graph path the change writes to. */
  readonly target: string;
  /** The ids of the item before the change: empty for a create. */
  readonly before: ItemIds;
};

/**
 * How to reverse one change: the `kind` of an undo handler (undo-run.ts) and
 * the data that handler needs, plain JSON. The journal keeps it as it is.
 */
export type UndoStep = { readonly kind: string; readonly data: Readonly<Record<string, unknown>> };

/** One change a command made: the ids before and after it, and how to reverse it. */
export type AppliedChange = PlannedChange & { readonly after: ItemIds; readonly undo: UndoStep };

/** What the writes of a plan give: the command's own answer, and each change made. */
export type AppliedWrite = { readonly value: unknown; readonly changes: ReadonlyArray<AppliedChange> };

/** The read members of the write graph: all a plan may call before the writes. */
export type GraphReads = Pick<WriteGraph, 'get' | 'getBinary' | 'fetchUrl'>;

/**
 * An item of a batch that the plan passes over: it needs no change, or it is
 * no longer there. The dry-run shows it with the reason.
 */
export type SkippedItem = { readonly target: string; readonly before: ItemIds; readonly reason: string };

/** The changes a command will make, and the writes that make them. */
export type WritePlan = {
  readonly changes: ReadonlyArray<PlannedChange>;
  /** The items of a batch that the plan passes over. The apply reports them again in its own answer. */
  readonly skipped?: ReadonlyArray<SkippedItem>;
  readonly apply: (graph: WriteGraph) => Promise<Result<AppliedWrite, GraphError>>;
};

/**
 * What a command exports as `plan`: it reads what it needs and says each
 * change. It gets the read members only, so a dry-run cannot write. `files`
 * reads a local file that a param names (read-local-file.ts).
 */
export type PlanWrite = (reads: GraphReads, params: Record<string, string>, files?: LocalFileReader) => Promise<Result<WritePlan, GraphError>>;

/** A plan bound to the graph of the run (command-graph.ts): its writes need no graph. */
export type BoundPlan = {
  readonly changes: ReadonlyArray<PlannedChange>;
  readonly skipped?: ReadonlyArray<SkippedItem>;
  readonly apply: () => Promise<Result<AppliedWrite, GraphError>>;
};

/** The plan as the shells see it, on the graph `G` they hold. */
export type BoundPlanWrite<G> = (graph: G, params: Record<string, string>, files?: LocalFileReader) => Promise<Result<BoundPlan, GraphError>>;
