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
  description: 'Widget CLI.',
  exampleCommand: 'list-widgets',
  binName: 'widget-cli',
  packageName: '@widgets/cli',
  // One bin for the three packages, as the single package has: no command lives elsewhere.
  bins: { auth: 'widget-cli', read: 'widget-cli', write: 'widget-cli' },
  ...overrides,
});

/**
 * The widget package as the read package of a split: auth and write are other
 * bins, so a command of theirs lives elsewhere. Its texts carry the bin
 * placeholders, so a renderer that does not fill them shows `{...bin}`.
 */
export const fakeReadPackage = (overrides: Partial<CommandPackage> = {}): CommandPackage =>
  fakeCommandPackage({
    registry: {
      'list-widgets': fakeCommand({
        summary: 'List the widgets. Page with `{read-bin} next-page`.',
        options: [
          { name: 'top', key: 'top', required: false, description: 'How many widgets; `{auth-bin} status` shows the token first.' },
          { name: 'shelf-id', key: 'shelfId', required: true, description: 'The shelf, as `{read-bin} list-shelves` returns it.' },
        ],
        example: '{read-bin} list-widgets --top 3',
      }),
      'get-widget-file': fakeCommand({ summary: 'Download a widget file. See `{bin} docs get-widget-file`.', producesBytes: true }),
    },
    lifecycle: [{ ...SIGN_IN_ENTRY, summary: 'Sign in to the widget service. Check with `{auth-bin} status`.', example: '{auth-bin} sign-in' }],
    bins: { auth: 'widget-auth', read: 'widget-cli', write: 'widget-write' },
    ...overrides,
  });
