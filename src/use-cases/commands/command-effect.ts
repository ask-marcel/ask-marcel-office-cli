import type { CommandEffect } from './command-types.ts';

type WriteEffect = Exclude<CommandEffect, 'read'>;

/**
 * What each write class does to the tenant, worded once, so a new write class
 * (package split, D9) is described here and nowhere else. The per-command docs
 * read it. `action` completes "<commands>: ...".
 */
const WRITE_EFFECTS: Readonly<Record<WriteEffect, { readonly action: string }>> = {
  draft: { action: 'create or update an UNSENT mail draft (this CLI cannot send mail)' },
  'transient-upload': {
    action:
      'upload the attachment to a temporary file in the `.ask-marcel-temp` folder of your OneDrive to convert it, then try to delete that file (a failed cleanup can leave it there)',
  },
};

export { WRITE_EFFECTS };
