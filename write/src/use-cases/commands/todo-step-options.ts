import { z } from 'zod';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

/* The options of the step commands: the step, its text and its check mark. */

export const checkedSchema = z.enum(['true', 'false']);

export const STEP_ID_OPTION: CommandOptionMeta = {
  name: 'step-id',
  key: 'stepId',
  required: true,
  description: 'The id of the step, as get-todo-task with --expand checklistItems returns it.',
  argumentHint: { kind: 'idOrName' },
};

export const STEP_TITLE_OPTION: CommandOptionMeta = { name: 'title', key: 'title', required: false, description: 'The text of the step.' };

export const CHECKED_OPTION: CommandOptionMeta = {
  name: 'checked',
  key: 'checked',
  required: false,
  description: '`true` checks the step off, `false` clears the check mark. A new step is not checked by default.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};
