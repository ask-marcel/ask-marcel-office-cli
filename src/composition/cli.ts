import { Command } from 'commander';
import type { AuthManager } from '../infra/auth.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { render, renderError } from '../presenter/output.ts';
import { buildManifest, renderSingleCommand } from '../use-cases/commands/docs.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, PAGINATION_HINT } from '../use-cases/commands/docs-render.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import * as login from '../use-cases/commands/login.ts';
import * as logout from '../use-cases/commands/logout.ts';
import * as update from '../use-cases/commands/update.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import { detectPackageManager } from './package-manager.ts';

type BuildCliDeps = {
  readonly auth: AuthManager;
  readonly graph: GraphClient;
  readonly logger: Logger;
  readonly processRunner: ProcessRunner;
  readonly version?: string;
  readonly packageManager?: 'npm' | 'bun';
  readonly onCommandError?: () => void;
};

const buildCli = (deps: BuildCliDeps): Command => {
  const { auth, graph, logger, processRunner, version } = deps;
  const program = new Command();

  const fail = (message: string): void => {
    renderError(message);
    deps.onCommandError?.();
  };

  program
    .name('ask-marcel')
    .description('Microsoft Graph CLI')
    .version(version ?? '0.0.0');

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

  const LIFECYCLE_COMMANDS: ReadonlySet<string> = new Set(['login', 'logout', 'update', 'docs', 'help', 'help-json']);

  const docsCmd = program
    .command('docs')
    .description('Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`).')
    .argument('<command>', 'Command name to show docs for (run `ask-marcel --help` to list every command).')
    .action((commandName: string) => {
      const result = renderSingleCommand(cmdRegistry, commandName);
      if (result.ok) {
        process.stdout.write(`${result.value}\n`);
        return;
      }
      if (LIFECYCLE_COMMANDS.has(commandName)) {
        const lifecycleCmd = program.commands.find((c) => c.name() === commandName);
        if (lifecycleCmd) {
          process.stdout.write(lifecycleCmd.helpInformation());
          return;
        }
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
        if (result.ok) render(result.value, logger);
        else fail(result.error.message);
      });
    }
  }

  return program;
};

export { buildCli };
export type { BuildCliDeps };
