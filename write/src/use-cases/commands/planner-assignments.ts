import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { ASSIGNMENT } from '../../domain/planner-fields.ts';
import type { PlannerPerson } from './planner-people.ts';
import { namesOf, readMe, resolvePeople, withoutConsent } from './planner-people.ts';

/*
 * The assignments of a create or an update of a task: the PATCH entries of
 * --assign and --unassign, the names that the dry-run shows, and the refusal
 * of a change for another person with no --notify-others true.
 */

/** The params of the people of a task. */
export type PeopleParams = { readonly assign?: string; readonly unassign?: string; readonly notifyOthers?: string };

/** The assignment entries of a PATCH or a POST, and the words of the dry-run. */
export type AssignmentChange = { readonly entries: Record<string, unknown>; readonly shown: Record<string, string> };

const NONE: AssignmentChange = { entries: {}, shown: {} };

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const hasId = (people: ReadonlyArray<PlannerPerson>, id: string): boolean => people.some((person) => person.id === id);

// The entries that change: an assignment to a person not assigned yet, the removal of a person assigned now.
const changeOf = (assign: ReadonlyArray<PlannerPerson>, unassign: ReadonlyArray<PlannerPerson>, current: Readonly<Record<string, unknown>>): AssignmentChange => {
  const added = assign.filter((person) => !Object.hasOwn(current, person.id));
  const removed = unassign.filter((person) => Object.hasOwn(current, person.id));
  return {
    entries: { ...Object.fromEntries(added.map((person) => [person.id, ASSIGNMENT])), ...Object.fromEntries(removed.map((person) => [person.id, null])) },
    shown: { ...(added.length === 0 ? {} : { assign: namesOf(added) }), ...(removed.length === 0 ? {} : { unassign: namesOf(removed) }) },
  };
};

/** The change of the assignments of `params` on a task whose assignments are `current` (empty for a new task). */
export const assignmentChangeOf = async (reads: GraphReads, params: PeopleParams, current: Readonly<Record<string, unknown>>): Promise<Result<AssignmentChange, GraphError>> => {
  if (params.assign === undefined && params.unassign === undefined) return ok(NONE);
  const me = await readMe(reads);
  if (!me.ok) return me;
  const assign = await resolvePeople(reads, me.value, params.assign ?? '', '--assign');
  if (!assign.ok) return assign;
  const unassign = await resolvePeople(reads, me.value, params.unassign ?? '', '--unassign');
  if (!unassign.ok) return unassign;
  if (assign.value.some((person) => hasId(unassign.value, person.id)))
    return refused('planner_assign_and_unassign', '--assign and --unassign name the same person: give the person in one of them.');
  const refusal = withoutConsent([...assign.value, ...unassign.value], params.notifyOthers, '--assign or --unassign');
  return refusal ?? ok(changeOf(assign.value, unassign.value, current));
};
