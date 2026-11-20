import type { Command, CommandEffect } from './command-types.ts';

type WriteEffect = Exclude<CommandEffect, 'read'>;

type WriteGroup = { readonly effect: WriteEffect; readonly names: ReadonlyArray<string> };

/**
 * What each write class does to the tenant, worded once. The MCP annotations,
 * the run-write-command and top-level `--help` sentences and the per-command
 * docs all read this table, so a new write class (package split, D9) is
 * described here and nowhere else. `action` completes "<commands>: ...".
 */
const WRITE_EFFECTS: Readonly<Record<WriteEffect, { readonly destructive: boolean; readonly action: string }>> = {
  draft: { destructive: false, action: 'create or update an UNSENT mail draft (this CLI cannot send mail)' },
  'transient-upload': {
    destructive: false,
    action:
      'upload the attachment to a temporary file in the `.ask-marcel-temp` folder of your OneDrive to convert it, then try to delete that file (a failed cleanup can leave it there)',
  },
  'mail-state': {
    destructive: false,
    action: 'change the state of a mail message: mark it read or unread, flag it, set its categories, or move it to another folder (nothing is sent)',
  },
  'task-write': {
    destructive: false,
    action:
      'create or change a To Do list, a To Do or Planner task, a Planner bucket or the names of a Planner plan, or a part of a task (a step, a checklist item, an attachment or a link, which a change can also remove). The file of a Planner task goes into the folder "Planner attachments" of your OneDrive',
  },
  // The one destructive class: MCP serves it on run-destructive-command, so the host sees the mark on these commands only.
  delete: {
    destructive: true,
    action:
      'delete or move an item: a mail draft goes to Deleted Items, from where the undo moves it back to Drafts. A To Do task, step or attachment, an attachment of a mail draft, a calendar draft, a Planner task, or an empty Planner bucket, is deleted (a To Do move deletes the original after its copy), and the undo makes it again from the run journal, with a new id (a calendar draft also gets a new Teams link). The delete of a Planner bucket with tasks has no undo',
  },
  'calendar-draft': {
    destructive: false,
    action: 'create or change a DRAFT event in your calendar: it has no attendees, so no invitation is sent (you add the attendees and send it yourself from Outlook)',
  },
};

/** The registry's write commands, one group per effect (in the order the registry first names it), names sorted. */
const writeGroups = (registry: Readonly<Record<string, Command<never>>>): ReadonlyArray<WriteGroup> => {
  const groups = new Map<WriteEffect, ReadonlyArray<string>>();
  for (const [name, command] of Object.entries(registry)) {
    const effect = command.meta.effect;
    if (effect !== 'read') groups.set(effect, [...(groups.get(effect) ?? []), name]);
  }
  return [...groups].map(([effect, names]) => ({ effect, names: names.toSorted((a, b) => a.localeCompare(b)) }));
};

/** One sentence per write class, e.g. `create-mail-draft, update-mail-draft: create or update an UNSENT mail draft (...)`. */
const describeWrites = (registry: Readonly<Record<string, Command<never>>>): string =>
  writeGroups(registry)
    .map((group) => `${group.names.join(', ')}: ${WRITE_EFFECTS[group.effect].action}`)
    .join('. ');

export { describeWrites, WRITE_EFFECTS, writeGroups };
