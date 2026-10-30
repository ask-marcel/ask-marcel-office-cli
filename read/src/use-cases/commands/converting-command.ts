import type { ReadCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { Converters } from '../ports/converters.ts';

/*
 * A read command that runs a document converter takes the Converters port when
 * the registry is built (ADR 0004): `createExecute` for a command on Graph,
 * `createExecuteLocal` for one on a local file. Every other command is a plain
 * ReadCommand.
 */
type ConvertingCommand = Omit<ReadCommand, 'execute'> & { readonly createExecute: (converters: Converters) => ReadCommand['execute'] };
type LocalConvertingCommand = Omit<ReadCommand, 'executeLocal'> & { readonly createExecuteLocal: (converters: Converters) => NonNullable<ReadCommand['executeLocal']> };
type CommandModule = ReadCommand | ConvertingCommand | LocalConvertingCommand;

/** The command, its converters in place. */
const withConverters = (module: CommandModule, converters: Converters): ReadCommand => {
  if ('createExecute' in module) {
    const { createExecute, ...rest } = module;
    return { ...rest, execute: createExecute(converters) };
  }
  if ('createExecuteLocal' in module) {
    const { createExecuteLocal, ...rest } = module;
    return { ...rest, executeLocal: createExecuteLocal(converters) };
  }
  return module;
};

export { withConverters };
export type { CommandModule, ConvertingCommand, LocalConvertingCommand };
