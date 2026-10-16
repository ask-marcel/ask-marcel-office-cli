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
};

/** The registry's write commands, one group per effect (in the order the registry first names it), names sorted. */
const writeGroups = (registry: Readonly<Record<string, Command>>): ReadonlyArray<WriteGroup> => {
  const groups = new Map<WriteEffect, ReadonlyArray<string>>();
  for (const [name, command] of Object.entries(registry)) {
    const effect = command.meta.effect;
    if (effect !== 'read') groups.set(effect, [...(groups.get(effect) ?? []), name]);
  }
  return [...groups].map(([effect, names]) => ({ effect, names: names.toSorted((a, b) => a.localeCompare(b)) }));
};

/** One sentence per write class, e.g. `create-mail-draft, update-mail-draft: create or update an UNSENT mail draft (...)`. */
const describeWrites = (registry: Readonly<Record<string, Command>>): string =>
  writeGroups(registry)
    .map((group) => `${group.names.join(', ')}: ${WRITE_EFFECTS[group.effect].action}`)
    .join('. ');

export { describeWrites, WRITE_EFFECTS, writeGroups };
