import { CONVERTERS } from '../infra/converters.ts';
import { createReadCommands } from '../use-cases/commands/index.ts';

/*
 * The read registry, built once for the composition root, the library entry and
 * the single package (ADR 0004): the one place that hands the commands their
 * converters.
 */
export const commands = createReadCommands(CONVERTERS);
