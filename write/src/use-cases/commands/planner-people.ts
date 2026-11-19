import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { isNotThere } from './mail-folder-reader.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The people of an assignment. Planner sends a notification to a person that
 * a task is assigned to, so an assignment to a person other than the user is
 * refused unless the command has `--notify-others true`. The same is true for
 * the removal of the assignment of another person: its undo assigns that
 * person again. `me` is the user; another person is a mail address or a user
 * id, checked before it goes into the path of the read.
 */

/** A person of an assignment: the user id, the name, and whether it is the user. */
export type PlannerPerson = { readonly id: string; readonly name: string; readonly isMe: boolean };

/** The option that permits an assignment to another person. */
export const NOTIFY_OTHERS = 'true';

const userSchema = z.object({ id: z.string().min(1), displayName: z.string().nullable().optional() }).loose();

const USER_ID = /^[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}$/i;
// A part of a mail address: no space, `@`, slash, `?`, `#` or `%`, so it cannot change the path of the read.
const MAIL_PART = /^[^\s@/\\?#%]+$/;

const isMail = (spec: string): boolean => {
  const parts = spec.split('@');
  return parts.length === 2 && parts.every((part) => MAIL_PART.test(part)) && (parts[1] ?? '').includes('.');
};

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

type User = { readonly id: string; readonly name: string };

const readUser = async (reads: GraphReads, path: string): Promise<Result<User, GraphError>> => {
  const read = await reads.get(`${path}?$select=id,displayName`);
  if (!read.ok) return read;
  const user = userSchema.safeParse(read.value);
  return user.success ? ok({ id: user.data.id, name: user.data.displayName ?? user.data.id }) : err(unreadable('a user'));
};

/** The user id of the user. */
export const readMe = async (reads: GraphReads): Promise<Result<User, GraphError>> => readUser(reads, '/me');

// The person of the item `index` of the option `what`.
const personOf = async (reads: GraphReads, me: User, spec: string, at: string): Promise<Result<PlannerPerson, GraphError>> => {
  if (spec.toLowerCase() === 'me') return ok({ ...me, isMe: true });
  if (!USER_ID.test(spec) && !isMail(spec)) return refused('planner_person_invalid', `${at} is not "me", a mail address or a user id.`);
  const user = await readUser(reads, `/users/${encodeURIComponent(spec)}`);
  if (!user.ok) return isNotThere(user.error) ? refused('planner_person_not_found', `${at}: no user of the tenant has this mail address or id.`) : user;
  return ok({ ...user.value, isMe: user.value.id === me.id });
};

/** The people of a list of `me`, mail addresses and user ids, separated by commas. `what` names the option in a refusal. */
export const resolvePeople = async (reads: GraphReads, me: User, raw: string, what: string): Promise<Result<ReadonlyArray<PlannerPerson>, GraphError>> => {
  const specs = raw
    .split(',')
    .map((spec) => spec.trim())
    .filter((spec) => spec !== '');
  const people: PlannerPerson[] = [];
  for (const [index, spec] of specs.entries()) {
    const person = await personOf(reads, me, spec, `${what}: item ${index + 1}`);
    if (!person.ok) return person;
    if (!people.some((known) => known.id === person.value.id)) people.push(person.value);
  }
  return ok(people);
};

/** The refusal of an assignment of another person with no `--notify-others true`, or undefined. */
export const withoutConsent = (people: ReadonlyArray<PlannerPerson>, notifyOthers: string | undefined, what: string): Result<never, GraphError> | undefined => {
  const others = people.filter((person) => !person.isMe).length;
  if (others === 0 || notifyOthers === NOTIFY_OTHERS) return undefined;
  return refused(
    'planner_notify_others_needed',
    `${what} names ${others} person(s) other than you. Planner sends each of them a notification, so this command needs --notify-others true for it.`
  );
};

/** The names of the people, for the dry-run: the user is "you". */
export const namesOf = (people: ReadonlyArray<PlannerPerson>): string => people.map((person) => (person.isMe ? `${person.name} (you)` : person.name)).join(', ');
