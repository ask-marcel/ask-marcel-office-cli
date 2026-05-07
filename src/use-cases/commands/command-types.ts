import type { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';

type CommandSchema = z.ZodType;
type CommandExecute = (graph: GraphClient, params: Record<string, string>) => Promise<Result<unknown, import('../../infra/graph-client.ts').GraphError>>;

type CommandCategory = 'auth' | 'drive' | 'excel' | 'sharepoint' | 'tasks' | 'mail' | 'notes' | 'user' | 'calendar' | 'contacts' | 'chats' | 'teams' | 'meta' | 'lifecycle';

type CommandHttpMethod = 'GET' | 'POST' | 'PATCH' | 'DELETE';

type CommandOptionAlias = {
  readonly name: string; // kebab-case (the user-facing flag, without `--`)
  readonly key: string; // camelCase (commander auto-keyed name)
};

type CommandOptionMeta = {
  readonly name: string;
  readonly key: string;
  readonly description: string;
  /**
   * `true` for required flags (the historical default; commander rejects the
   * invocation if the flag is missing). `false` for optional flags such as the
   * OData passthrough query parameters (`--top`, `--filter`, …) which
   * commands accept but do not demand.
   */
  readonly required: boolean;
  /**
   * Optional secondary spellings of the same flag. Both the canonical
   * `name` and every alias name are accepted on the command line; values
   * passed under an alias are normalized to the canonical `key` before
   * the schema runs. The canonical name is what `--help` shows first.
   */
  readonly aliases?: ReadonlyArray<CommandOptionAlias>;
};

type CommandMeta = {
  readonly summary: string;
  readonly category: CommandCategory;
  readonly graphMethod: CommandHttpMethod;
  readonly graphPathTemplate: string;
  readonly graphDocsUrl: string;
  readonly options: ReadonlyArray<CommandOptionMeta>;
  readonly example: string;
  readonly responseShape?: string;
  readonly bodyTemplate?: string;
  readonly pagination?: true;
};

type Command = {
  readonly schema: CommandSchema;
  readonly execute: CommandExecute;
  readonly meta: CommandMeta;
};

export type { Command, CommandCategory, CommandExecute, CommandHttpMethod, CommandMeta, CommandOptionAlias, CommandOptionMeta, CommandSchema };
