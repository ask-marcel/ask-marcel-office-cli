/*
 * A small package of made-up commands for the tests of the registry-agnostic
 * shells (CLI builder, MCP gateway, run-registry-command, docs builders).
 *
 * The names are deliberately foreign to this CLI (`widget-cli`, `list-widgets`),
 * so a shell that still reads the real registry, the real lifecycle entries or
 * the `ask-marcel-office` names shows up as a wrong word in the output.
 */
import { z } from 'zod';
import { ok } from '../domain/result.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import type { Command, CommandMeta } from '../use-cases/commands/command-types.ts';
import type { CommandManifestEntry } from '../use-cases/commands/docs-render.ts';

/** A command that answers with the params it was given, so a test sees that it ran and with what. */
export const fakeCommand = (meta: Partial<CommandMeta> = {}): Command => ({
  schema: z.object({}).strict(),
  execute: async (_graph, params) => ok({ ranWith: params }),
  meta: {
    summary: 'Do a widget thing. Second sentence.',
    category: 'drive',
    graphMethod: 'GET',
    graphPathTemplate: '/widgets',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/widgets',
    options: [{ name: 'top', key: 'top', required: false, description: 'How many widgets.' }],
    example: 'widget-cli list-widgets --top 3',
    effect: 'read',
    ...meta,
  },
});

export const SIGN_IN_ENTRY: CommandManifestEntry = {
  name: 'sign-in',
  summary: 'Sign in to the widget service. It opens no browser.',
  category: 'lifecycle',
  graphMethod: 'GET',
  graphPathTemplate: '(lifecycle) signs in to the widget service; not a Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/auth-v2-user',
  options: [],
  example: 'widget-cli sign-in',
};

/**
 * Three GET reads (one returns bytes, one returns media), one read-only POST
 * search and one draft write, plus the `sign-in` lifecycle entry.
 */
export const fakeCommandPackage = (overrides: Partial<CommandPackage> = {}): CommandPackage => ({
  registry: {
    'list-widgets': fakeCommand(),
    'get-widget-file': fakeCommand({ summary: 'Download a widget file.', producesBytes: true }),
    'get-widget-images': fakeCommand({ summary: 'Extract the widget images.', producesMedia: true }),
    'find-widgets': fakeCommand({ summary: 'Search the widgets.', category: 'meta', graphMethod: 'POST' }),
    'save-widget-draft': fakeCommand({ summary: 'Save a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'draft' }),
  },
  lifecycle: [SIGN_IN_ENTRY],
  binName: 'widget-cli',
  packageName: '@widgets/cli',
  ...overrides,
});
