import { Command } from 'commander';
import type { AuthManager } from '../infra/auth.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { render, renderError } from '../presenter/output.ts';
import { buildManifest, renderSingleCommand } from '../use-cases/commands/docs.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, PAGINATION_HINT } from '../use-cases/commands/docs-render.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import * as login from '../use-cases/commands/login.ts';
import * as logout from '../use-cases/commands/logout.ts';
import { persistIfRequested } from '../use-cases/commands/output-path.ts';
import * as update from '../use-cases/commands/update.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import { detectPackageManager } from './package-manager.ts';

type BuildCliDeps = {
  readonly auth: AuthManager;
  readonly graph: GraphClient;
  readonly logger: Logger;
  readonly processRunner: ProcessRunner;
  readonly fs: FileSystem;
  readonly version?: string;
  readonly packageManager?: 'npm' | 'bun';
  readonly onCommandError?: () => void;
};

const buildCli = (deps: BuildCliDeps): Command => {
  const { auth, graph, logger, processRunner, fs, version } = deps;
  const program = new Command();

  const fail = (message: string): void => {
    renderError(message);
    deps.onCommandError?.();
  };

  // Single-stream JSON contract: commander's parser errors (unknown option,
  // missing required, unknown command, etc.) used to land on stderr as plain
  // text, while validation and Graph errors landed on stdout as JSON. An LLM
  // capturing only stdout silently lost the parser cases. Suppress commander's
  // stderr writer and intercept its CommanderError in exitOverride so we can
  // render the same JSON envelope every other path uses.
  program.configureOutput({
    writeErr: () => undefined,
  });
  program.exitOverride((err) => {
    if (err.code === 'commander.helpDisplayed' || err.code === 'commander.version' || err.code === 'commander.help') return;
    // Commander prefixes its messages with `error: ` (e.g. "error: unknown option '--foo'"),
    // but the JSON envelope's outer `ok: false` already conveys errorness — strip the
    // redundant prefix so consumers don't see `{"ok":false,"error":"error: ..."}`.
    const stripped = err.message.startsWith('error: ') ? err.message.slice('error: '.length) : err.message;
    fail(stripped);
    throw err;
  });

  program
    .name('ask-marcel')
    .description('Microsoft Graph CLI')
    .version(version ?? '0.0.0')
    .option(
      '--output-path <path>',
      'Globally available. When the command returns inlined bytes (`{contentType, size, base64}` for binary or `{..., text}` for text), decode and write them to <path>, replacing the inline field with `savedTo: <path>` in the JSON envelope. Use this for multi-MB PDFs / images so the LLM never has to round-trip a base64 string through stdout. Parent directories are auto-created. When applied to a command whose response has neither `base64` nor `text` (e.g. plain JSON gets like `get-current-user`) the CLI emits a clear `{"ok":false,"error":"--output-path: <cmd> did not return inlined bytes …"}` envelope rather than silently writing nothing — a JSON-only command paired with this flag is almost certainly a mistake.'
    );

  // Audit v1.0.0 §2.3: bare `ask-marcel` (no subcommand) used to silently
  // exit 1 with zero output. We intercept that case BEFORE Commander parses
  // so we don't break the existing `unknown subcommand` error path. Hooked on
  // `preAction` of every subcommand would be wrong (it never fires for the
  // bare case); instead we override `parseAsync` itself.
  const originalParseAsync = program.parseAsync.bind(program);
  program.parseAsync = (async (argv?: readonly string[], options?: { readonly from?: 'node' | 'electron' | 'user' }) => {
    const args = argv ?? process.argv;
    const from = options?.from ?? 'node';
    const userArgsStart = from === 'node' || from === 'electron' ? 2 : 0;
    if (args.length <= userArgsStart) {
      program.outputHelp();
      return program;
    }
    const fixedOptions = options === undefined ? undefined : ({ from } as { from: 'node' | 'electron' | 'user' });
    return originalParseAsync(args as string[], fixedOptions);
  }) as typeof program.parseAsync;

  // Override Commander's built-in `help <command>` (which silently exits 1 on
  // unknown subcommands — audit v1.0.0 §1.2). Disable the built-in first, then
  // register our own with the same JSON-envelope contract every other path uses.
  program.helpCommand(false);
  program
    .command('help [command]', { hidden: true })
    .description('Show docs for a command (alias of `docs <command>`). Without an argument, prints the global `--help` text.')
    .action((commandName?: string) => {
      if (commandName === undefined) {
        program.outputHelp();
        return;
      }
      const result = renderSingleCommand(cmdRegistry, commandName);
      if (result.ok) {
        process.stdout.write(`${result.value}\n`);
        return;
      }
      fail(`Unknown command "${result.error.name}". Run \`ask-marcel --help\` to list every command.`);
    });

  program
    .command('help-json')
    .description(
      'Print the full machine-readable command manifest as JSON to stdout (same content as `docs/commands.json`). Token-friendly alternative to `--help` for LLM consumers.'
    )
    .action(() => {
      process.stdout.write(`${JSON.stringify(buildManifest(cmdRegistry, 'ask-marcel-office-cli', version ?? '0.0.0'))}\n`);
    });

  program.commandsGroup('Lifecycle:');

  const loginCmd = program
    .command('login')
    .description('Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback).')
    .action(async () => {
      const result = await login.execute(auth);
      if (result.ok) render({ status: 'authenticated' }, logger);
      else fail(result.error.type === 'auth_cancelled' ? 'Authentication cancelled' : result.error.message);
    });
  loginCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel login',
      'Token cache:   ~/.ask-marcel/token-cache.json (access + refresh tokens, JSON, 0600).',
      'Browser data:  ~/.ask-marcel/browser-profile/ (Playwright persistent context).',
      'Scopes:        granted by Microsoft to the Teams web client (CLIENT_ID 5e3ce6c0-...);',
      '               this CLI cannot request additional scopes. To inspect the granted set,',
      '               run `ask-marcel scopes-check`.',
      'Stuck flow:    `ask-marcel logout` then re-run; the browser fallback opens a fresh Edge / Chrome window.',
    ].join('\n  ')
  );

  const logoutCmd = program
    .command('logout')
    .description('Clear the cached Microsoft Graph token so the next command forces a fresh sign-in.')
    .action(async () => {
      const result = await logout.execute(auth);
      if (result.ok) render({ status: 'logged_out' }, logger);
      else fail(result.error.type === 'auth_cancelled' ? 'Logout cancelled' : result.error.message);
    });
  logoutCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel logout',
      'Removes:       ~/.ask-marcel/token-cache.json (access + refresh tokens).',
      'Leaves alone:  ~/.ask-marcel/browser-profile/ (delete it manually if you want a clean Playwright session too).',
      'Verify clean:  ls ~/.ask-marcel/  (token-cache.json should be gone).',
    ].join('\n  ')
  );

  const updateCmd = program
    .command('update')
    .description('Re-install the latest published ask-marcel from npm, in place. Auto-detects whether you originally installed via npm or bun.')
    .action(async () => {
      const manager = deps.packageManager ?? detectPackageManager(process.argv[1] ?? '');
      const result = await update.execute(processRunner, manager);
      if (result.ok) render({ status: 'updated', via: manager }, logger);
      else if (result.error.type === 'spawn_failed') fail(`update failed: ${result.error.message}`);
      else fail(`update install exited with code ${result.error.exitCode}`);
    });
  updateCmd.addHelpText(
    'after',
    [
      '',
      'Example:      ask-marcel update',
      'Detection:    based on the bin path of the running CLI.',
      '              `/usr/local/lib/node_modules/...` -> npm, `~/.bun/install/...` -> bun.',
      'Side effect:  shells out to `npm i -g ask-marcel-office-cli@latest` or `bun add -g ...`.',
      'Token cache:  preserved (this only re-installs the JS bundle).',
      'Local clone:  do NOT use `update` — pull and re-run `bun install` instead.',
    ].join('\n  ')
  );

  const docsCmd = program
    .command('docs')
    .description(
      'Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`). Lifecycle commands (login/logout/update/docs/help-json) are also covered — they ship as manifest entries under category `lifecycle`.'
    )
    .argument('<command>', 'Command name to show docs for (run `ask-marcel --help` to list every command).')
    .action((commandName: string) => {
      const result = renderSingleCommand(cmdRegistry, commandName);
      if (result.ok) {
        process.stdout.write(`${result.value}\n`);
        return;
      }
      fail(`Unknown command "${result.error.name}". Run \`ask-marcel --help\` to list every command.`);
    });
  docsCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel docs list-mail-messages',
      'Lifecycle:     `ask-marcel docs login` (or logout / update / docs) prints the same --help that command would, so you can introspect lifecycle commands the same way.',
    ].join('\n  ')
  );

  for (const category of CATEGORY_ORDER) {
    const entries = Object.entries(cmdRegistry).filter(([, c]) => c.meta.category === category);
    if (entries.length === 0) continue;
    program.commandsGroup(`${CATEGORY_LABELS[category]}:`);
    for (const [name, cmd] of entries) {
      const commandDef = program.command(name).description(cmd.meta.summary);
      for (const opt of cmd.meta.options) {
        if (opt.aliases && opt.aliases.length > 0) {
          // When aliases exist we can't use `requiredOption` for the canonical
          // — Commander would reject alias-only invocations (the canonical
          // long flag would be missing). Schema validation
          // (z.string().min(1)) enforces required-ness instead.
          commandDef.option(`--${opt.name} <value>`, opt.description);
          for (const alias of opt.aliases) {
            commandDef.option(`--${alias.name} <value>`, `(alias for --${opt.name})`);
          }
        } else if (opt.required) {
          commandDef.requiredOption(`--${opt.name} <value>`, opt.description);
        } else {
          commandDef.option(`--${opt.name} <value>`, opt.description);
        }
      }
      const helpLines = [
        `\nGraph endpoint: ${cmd.meta.graphMethod} ${cmd.meta.graphPathTemplate}`,
        `Microsoft Learn: ${cmd.meta.graphDocsUrl}`,
        ...(cmd.meta.pagination ? [`\nPagination: ${PAGINATION_HINT}`] : []),
        ...(cmd.meta.bodyTemplate ? [`\nRequest body:\n  ${cmd.meta.bodyTemplate}`] : []),
        `\nExample:\n  ${cmd.meta.example}`,
      ];
      commandDef.addHelpText('after', helpLines.join('\n'));
      commandDef.action(async (opts: Record<string, string>) => {
        const normalized: Record<string, string> = { ...opts };
        for (const opt of cmd.meta.options) {
          for (const alias of opt.aliases ?? []) {
            const aliasValue = normalized[alias.key];
            if (typeof aliasValue === 'string' && normalized[opt.key] === undefined) {
              normalized[opt.key] = aliasValue;
            }
          }
        }
        const result = await cmd.execute(graph, normalized);
        if (!result.ok) {
          fail(result.error.message);
          return;
        }
        const outputPath = program.opts<{ outputPath?: string }>().outputPath;
        const persisted = await persistIfRequested(fs, outputPath, result.value);
        if (persisted.ok) {
          render(persisted.value, logger);
          return;
        }
        const failMessage = ((): string => {
          if (persisted.error.type === 'no_inlined_bytes')
            return `--output-path: ${name} did not return inlined bytes — this flag works with the families that produce a body to write: download-drive-item-as-pdf / -as-markdown, download-drive-item-version-as-pdf / -as-markdown / -content, download-onedrive-file-content, convert-mail-attachment-to-pdf / -to-markdown, convert-mail-to-markdown, get-mail-message-mime, get-my-profile-photo, get-onenote-page-content, get-sharepoint-site-onenote-page-content. Plain JSON commands (list-*, get-*-user, get-organization, list-recent-files, etc.) don't have a body to write — drop the flag for those.`;
          if (persisted.error.type === 'empty_path') return '--output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)';
          return `--output-path: write failed: ${persisted.error.message}`;
        })();
        fail(failMessage);
      });
    }
  }

  return program;
};

export { buildCli };
export type { BuildCliDeps };
