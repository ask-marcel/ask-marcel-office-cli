import { Command, InvalidArgumentError, Option } from 'commander';
import type { AuthManager } from '../infra/auth.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import type { ErrorSource } from '../presenter/error-hints.ts';
import type { OutputFormat } from '../presenter/output.ts';
import { render, renderError } from '../presenter/output.ts';
import type { RenderContext } from '../presenter/render-to-string.ts';
import { buildManifest, buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, paginationHintFor } from '../use-cases/commands/docs-render.ts';
import { firstSentence } from '../use-cases/commands/first-sentence.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import * as login from '../use-cases/commands/login.ts';
import { buildLoginSummary } from '../use-cases/commands/login-status.ts';
import * as logout from '../use-cases/commands/logout.ts';
import { persistIfRequested } from '../use-cases/commands/output-path.ts';
import * as update from '../use-cases/commands/update.ts';
import { buildRenderContext, formatOutputPathError, runRegistryCommand } from './run-registry-command.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import type { LoginAuthFactory } from './build-deps.ts';
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
  /**
   * Builds the AuthManager for an interactive `login` run (it may recapture
   * secondary tokens via the browser, unlike the command-path `auth`).
   * Supplied by the composition root (`build-deps.ts`); when omitted (tests),
   * the login action falls back to the injected `auth`, so a single fake
   * drives both `login.execute` and `getLastElevatedOutcome`.
   */
  readonly makeLoginAuth?: LoginAuthFactory;
};

const buildCli = (deps: BuildCliDeps): Command => {
  const { auth, graph, logger, processRunner, fs, version } = deps;
  const program = new Command();

  const getFormat = (): OutputFormat => {
    const raw = program.opts<{ output?: string }>().output;
    return raw === 'json' ? 'json' : 'text';
  };
  const renderOut = (data: unknown, sizeHintContext?: RenderContext): void => render(data, logger, getFormat(), sizeHintContext);
  const fail = (message: string, code?: string, source?: ErrorSource, retryAfterSeconds?: number): void => {
    renderError(message, getFormat(), code, source, retryAfterSeconds);
    deps.onCommandError?.();
  };

  // route help-json / docs through persistIfRequested so --output-path is honoured (was silently ignored).
  const writeOrPrintText = async (textBody: string, contentType: string, commandName: string): Promise<void> => {
    const outputPath = program.opts<{ outputPath?: string }>().outputPath;
    if (outputPath === undefined) {
      process.stdout.write(`${textBody}\n`);
      return;
    }
    const persisted = await persistIfRequested(fs, outputPath, { contentType, size: textBody.length, text: textBody });
    if (persisted.ok) {
      renderOut(persisted.value);
      return;
    }
    fail(formatOutputPathError(persisted.error, commandName, 'cli'));
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
    // pass Commander's structured code
    // (`commander.unknownOption`, `commander.missingMandatoryOptionValue`,
    // etc.) through to the error renderer so the hint table can match it and
    // surface `hint:` / `source:` for the CLI-input failure cases too. Prior
    // behaviour dropped the code, leaving Commander errors as bare
    // `{ok:false,error}` envelopes — inconsistent with the Graph error shape.
    // also stamp `source: 'cli'` explicitly
    // so the envelope shape stays stable even if a future Commander error
    // code variant doesn't have a matching rule yet.
    fail(stripped, err.code, 'cli');
    throw err;
  });

  // previous default `--help` ran ~60 KB because the
  // top-level subcommand listing rendered each command's full summary (often
  // 2-3 sentences). Default `--help` now truncates each subcommand description
  // to its first sentence in the top-level listing (~34 KB total). Per-command
  // `ask-marcel-office <cmd> --help` is never compacted — it always shows the full
  // description plus the `addHelpText` block, and `help-json` ships the full
  // summary unchanged. v1.4.0 surface-consolidation: the `--verbose` opt-out
  // was dropped (it was a one-trick toggle on this top-level listing only;
  // `help-json --terse` covers the same need with a structured payload).

  // Compact a subcommand's description to its first sentence — shared with the
  // `help-json --terse` projection via `firstSentence` so both stay consistent.

  // QA fix: derive the advertised endpoint counts from the registry so the
  // top-level description can never drift from the manifest again. The
  // hardcoded "164 GET + 1 POST" literal had gone stale (real: 169 GET + 2
  // POST) when search-all-accessible-sites was added without updating it.
  const getEndpointCount = Object.values(cmdRegistry).filter((c) => c.meta.graphMethod === 'GET').length;
  // 2026-06-15 (F-03), extended 2026-07-04: the only write commands are the three mail-draft ones
  // (create=POST, update=PATCH); everything else is a read or a search. Derive
  // BOTH lists from the manifest's `mutates` flag so the read-only narrative can
  // never drift from the registry. The previous code took every POST command and
  // labelled it "— searches, not mutations", which silently mislabelled
  // create-mail-draft as a search the moment it shipped (and omitted the PATCH
  // update-mail-draft entirely).
  const mutatingCommandNames = Object.entries(cmdRegistry)
    .filter(([, c]) => c.meta.mutates === true)
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));
  const searchPostNames = Object.entries(cmdRegistry)
    .filter(([, c]) => c.meta.graphMethod === 'POST' && c.meta.mutates !== true)
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));
  const surfaceDescription = `Microsoft Graph CLI. Read-mostly by design — the ONLY writes are the ${mutatingCommandNames.length} mail-draft commands (${mutatingCommandNames.join(', ')}), which can only create or update an UNSENT draft; the CLI cannot send mail, create or modify calendar items, or write files (there is no send-mail / send-draft / create-event / upload-file command). ${getEndpointCount} GET endpoints + ${searchPostNames.length} search POST (${searchPostNames.join(', ')}). Safe default for LLM autonomy.`;

  program
    .name('ask-marcel-office')
    .description(surfaceDescription)
    .version(version ?? '0.0.0')
    // override the help-formatter's subcommand
    // description renderer to compact long summaries down to their first
    // sentence in the TOP-LEVEL `ask-marcel-office --help` listing. Per-subcommand
    // `ask-marcel-office <cmd> --help` is untouched — Commander only consults
    // `subcommandDescription` when formatting parent's child list.
    .configureHelp({
      subcommandDescription: (cmd) => firstSentence(cmd.description()),
    })
    .option(
      '--output-path <path>',
      'Globally available. When the command returns inlined bytes (`{contentType, size, base64}` for binary or `{..., text}` for text), decode and write them to <path>, replacing the inline field with `savedTo: <path>` in the JSON envelope. Use this for multi-MB PDFs / images so the LLM never has to round-trip a base64 string through stdout. Parent directories are auto-created. When applied to a command whose response has neither `base64` nor `text` (e.g. plain JSON gets like `get-current-user`) the CLI emits a clear `{"ok":false,"error":"--output-path: <cmd> did not return inlined bytes …"}` envelope rather than silently writing nothing — a JSON-only command paired with this flag is almost certainly a mistake.'
    )
    .option(
      '--output-dir <dir>',
      'Globally available. For commands that return a `media` array (the `extract-*-images` family), decode and write every image to <dir>/<filename>, replacing each `base64` with `savedTo` in the JSON envelope. The directory is auto-created. Use this for image-heavy decks so the LLM never round-trips a base64 blob through stdout. Applied to a command that returns no media array, the CLI emits a clear error rather than writing nothing.'
    )
    .addOption(
      ((): Option => {
        // reject duplicate `--output` flags AND
        // validate the choice. Adding an argParser bypasses Commander's
        // built-in `.choices()` enforcement, so the parser has to do both
        // jobs. Track invocations via a closure flag (Commander seeds
        // `previous` from the default, which would otherwise make any single
        // `--output json` look like the second occurrence).
        const ALLOWED: ReadonlyArray<string> = ['text', 'json'];
        let outputSeen = false;
        return new Option(
          '--output <format>',
          'Output format. `text` (default, LLM-readable YAML-ish lines, ~30-60% fewer tokens on listings; errors render as `error: <message>`). `json` preserves the `{ok, data, nextLink?, deltaLink?, count?}` envelope for tool-chaining where unambiguous field extraction matters.'
        )
          .default('text')
          .argParser((value: string, previous: unknown): string => {
            if (outputSeen) {
              throw new InvalidArgumentError(`--output cannot be passed more than once (previous: "${String(previous)}", new: "${value}").`);
            }
            if (!ALLOWED.includes(value)) {
              throw new InvalidArgumentError(`Allowed choices are ${ALLOWED.join(', ')}.`);
            }
            outputSeen = true;
            return value;
          });
      })()
    );

  // explicit pointers to per-command help and the
  // machine-readable manifest. Without this, an LLM that hits the compact
  // top-level help has no in-band signal that `help-json` even exists.
  program.addHelpText(
    'after',
    [
      '',
      'For full per-command help:   ask-marcel-office <command> --help',
      'For machine-readable docs:   ask-marcel-office help-json [--terse] [--category mail]',
      'For per-command Markdown:    ask-marcel-office docs <command>',
      // `mcp` is intercepted in main.ts, not registered as a subcommand (that
      // would pull the SDK into every command's module graph), so Commander
      // cannot list it by itself. Name it here or it is undiscoverable.
      'To serve this over MCP:      ask-marcel-office mcp   (stdio; see README for `claude mcp add`)',
    ].join('\n  ')
  );

  // bare `ask-marcel-office` (no subcommand) used to silently
  // exit 1 with zero output. We intercept that case BEFORE Commander parses
  // so we don't break the existing `unknown subcommand` error path. Hooked on
  // `preAction` of every subcommand would be wrong (it never fires for the
  // bare case); instead we override `parseAsync` itself.
  const originalParseAsync = program.parseAsync.bind(program);
  program.parseAsync = async (argv?: readonly string[], options?: { readonly from?: 'node' | 'electron' | 'user' }) => {
    const args = argv ?? process.argv;
    const from = options?.from ?? 'node';
    const userArgsStart = from === 'node' || from === 'electron' ? 2 : 0;
    if (args.length <= userArgsStart) {
      program.outputHelp();
      return program;
    }
    const fixedOptions = options === undefined ? undefined : { from };
    return originalParseAsync(args, fixedOptions);
  };

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
      // pass an explicit `cli_unknown_command`
      // code so error-hints can match it structurally — gives the same
      // envelope shape as Commander's own `commander.unknownCommand` path.
      fail(`Unknown command "${result.error.name}". Run \`ask-marcel-office --help\` to list every command.`, 'cli_unknown_command');
    });

  program
    .command('help-json')
    .description(
      'Print the machine-readable command manifest as JSON. **Use `--terse --category <name>` for fresh-session discovery** — that combo is the actual token-friendly path (~8 KB for one category, vs ~505 KB unfiltered). The unflagged form is the *full* reference (every option / example / response shape per command) and is well over 10× the size of `ask-marcel-office --help`; reach for it only after `--terse` has narrowed the search. `--terse` alone projects to `{name, summary, category}` with each summary compacted to its first sentence (~33 KB across all categories). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.'
    )
    .option(
      '--terse',
      'Strip per-command options/example/graphPathTemplate/responseShape/etc., emitting only `{ name, summary, category }`. Roughly 95% smaller than the full manifest — use for command discovery, switch to the full manifest once you know which command to invoke.'
    )
    .option(
      '--category <name>',
      'Filter the manifest to a single category (one of: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta). Composes with `--terse`. Unknown categories return a structured `{ ok: false, error }` envelope.'
    )
    .action(async (opts: { readonly terse?: boolean; readonly category?: string }) => {
      // --output text was silently honored on
      // help-json, contradicting the global flag's contract. The manifest IS
      // JSON; serializing as text has no use case. Reject only when the user
      // EXPLICITLY passed `--output text` (defaulting to text and getting
      // JSON anyway is fine — that's the historical behavior).
      const outputSource = program.getOptionValueSource('output');
      if (outputSource === 'cli' && getFormat() === 'text') {
        fail(
          "help-json always emits JSON (that's the contract — the manifest is the LLM-consumable serialized form of every command's meta). Drop `--output text` for this command. To browse the manifest as Markdown, use `ask-marcel-office docs <command>` instead."
        );
        return;
      }
      // --terse projects to `{name, summary, category}`;
      // --category filters to a single category (validated against
      // CATEGORY_LABELS). Both flags compose.
      const fullOrTerse =
        opts.terse === true
          ? buildTerseManifest(cmdRegistry, 'ask-marcel-office-cli', version ?? '0.0.0')
          : buildManifest(cmdRegistry, 'ask-marcel-office-cli', version ?? '0.0.0');
      if (opts.category !== undefined) {
        const filtered = filterManifestByCategory(fullOrTerse, opts.category);
        if (!filtered.ok) {
          fail(`Unknown --category "${filtered.error.category}". Available categories: ${filtered.error.available.join(', ')}.`, 'cli_unknown_category');
          return;
        }
        await writeOrPrintText(JSON.stringify(filtered.value), 'application/json', 'help-json');
        return;
      }
      await writeOrPrintText(JSON.stringify(fullOrTerse), 'application/json', 'help-json');
    });

  program.commandsGroup('Lifecycle:');

  const loginCmd = program
    .command('login')
    .description(
      'Authenticate against Microsoft Graph via the Teams web client (cached token → refresh → browser). Already signed in? Reports all four cached tokens (basic / elevated / chatsvcagg / ic3) with their time-left and refresh route; --force re-captures every token via the browser.'
    )
    .option(
      '--force',
      'Ignore the cache and re-capture every token via the browser. Rarely needed, and NOT free: it clears the browser session first so the token grant re-fires, which deletes the 90-day "Stay signed in" cookie and usually means entering your password again. A plain `login` already re-captures the elevated (M365) token when it is missing and leaves that session intact, so reach for `--force` only when a tier is stuck and a plain `login` has not fixed it.'
    )
    .action(async () => {
      const force = loginCmd.opts<{ force?: boolean }>().force ?? false;
      // The composition root supplies `makeLoginAuth` (a login-configured
      // manager that may recapture secondary tokens via the browser); tests
      // omit it and the injected `auth` fake drives both execute and the
      // token-status read.
      const loginAuth = deps.makeLoginAuth ? deps.makeLoginAuth() : auth;
      const result = await login.execute(loginAuth, { force });
      if (!result.ok) {
        fail(result.error.type === 'auth_cancelled' ? 'Authentication cancelled' : result.error.message);
        return;
      }
      // Slim confirmation: which tokens are available now, and where to look next.
      // The decode-only read never opens a browser. Full per-token scopes + expiry
      // live in `scopes-check`; login just points there (and to `login --force`).
      const info = await graph.getCachedTokenInfo();
      if (!info.ok) {
        fail(info.error.message);
        return;
      }
      renderOut(
        buildLoginSummary({
          elevatedAvailable: info.value.elevated.available,
          chatsvcaggAvailable: info.value.chatsvcagg.available,
          ic3Available: info.value.ic3.available,
        })
      );
    });
  loginCmd.addHelpText(
    'after',
    [
      '',
      'Examples:      ask-marcel-office login          (sign in the first time, or show all four token statuses if already signed in)',
      '               ask-marcel-office login --force  (re-capture every token via the browser, ignoring the cache)',
      'Token cache:   ~/.ask-marcel/token-cache.json (access + refresh tokens, JSON, 0600).',
      'Browser data:  ~/.ask-marcel/browser-profile/ (Playwright persistent context).',
      'Scopes:        granted by Microsoft to the Teams web client (CLIENT_ID 5e3ce6c0-...);',
      '               this CLI cannot request additional scopes. To inspect the granted set,',
      '               run `ask-marcel-office scopes-check`.',
      'Stuck flow:    `ask-marcel-office logout` then re-run; the browser fallback opens a fresh Edge / Chrome window.',
    ].join('\n  ')
  );

  const logoutCmd = program
    .command('logout')
    .description('Clear the cached Microsoft Graph token so the next command forces a fresh sign-in.')
    .action(async () => {
      const result = await logout.execute(auth);
      if (result.ok) renderOut({ status: 'logged_out' });
      else fail(result.error.type === 'auth_cancelled' ? 'Logout cancelled' : result.error.message);
    });
  logoutCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel-office logout',
      'Removes:       ~/.ask-marcel/token-cache.json (access + refresh tokens).',
      'Leaves alone:  ~/.ask-marcel/browser-profile/ (delete it manually if you want a clean Playwright session too).',
      'Verify clean:  ls ~/.ask-marcel/  (token-cache.json should be gone).',
    ].join('\n  ')
  );

  const updateCmd = program
    .command('update')
    .description('Re-install the latest published ask-marcel-office from npm, in place. Auto-detects whether you originally installed via npm or bun.')
    .action(async () => {
      const manager = deps.packageManager ?? detectPackageManager(process.argv[1] ?? '');
      const result = await update.execute(processRunner, manager);
      if (result.ok) renderOut({ status: 'updated', via: manager });
      else if (result.error.type === 'spawn_failed') fail(`update failed: ${result.error.message}`);
      else fail(`update install exited with code ${result.error.exitCode}`);
    });
  updateCmd.addHelpText(
    'after',
    [
      '',
      'Example:      ask-marcel-office update',
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
    .argument('<command>', 'Command name to show docs for (run `ask-marcel-office --help` to list every command).')
    .action(async (commandName: string) => {
      const result = renderSingleCommand(cmdRegistry, commandName);
      if (!result.ok) {
        // structured `cli_unknown_command`
        // code so the envelope matches the `help <unknown>` and
        // `commander.unknownCommand` paths — single branch for LLM consumers.
        fail(`Unknown command "${result.error.name}". Run \`ask-marcel-office --help\` to list every command.`, 'cli_unknown_command');
        return;
      }
      await writeOrPrintText(result.value, 'text/markdown', 'docs');
    });
  docsCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel-office docs list-mail-messages',
      'Lifecycle:     `ask-marcel-office docs login` (or logout / update / docs) prints the same --help that command would, so you can introspect lifecycle commands the same way.',
    ].join('\n  ')
  );

  for (const category of CATEGORY_ORDER) {
    const entries = Object.entries(cmdRegistry).filter(([, c]) => c.meta.category === category);
    if (entries.length === 0) continue;
    program.commandsGroup(`${CATEGORY_LABELS[category]}:`);
    for (const [name, cmd] of entries) {
      const commandDef = program.command(name).description(cmd.meta.summary);
      // Audit round-7 B6: every single-value flag rejects repeated occurrences.
      // Commander.js by default last-wins on `--filter A --filter B` — surprising
      // for an LLM consumer that constructed two filters expecting both to apply.
      const noRepeatParser =
        (flagName: string) =>
        (value: string, previous: unknown): string => {
          if (typeof previous === 'string') {
            throw new InvalidArgumentError(
              `--${flagName} cannot be passed more than once (previous: "${previous}", new: "${value}"). Single-value flags reject duplicate occurrences.`
            );
          }
          return value;
        };
      // 2026-07-24: one name per flag. The alias registration branch that used
      // to live here (and forced plain `option` + schema-enforced requiredness
      // on aliased canonicals) went with the alias system.
      for (const opt of cmd.meta.options) {
        if (opt.required) {
          commandDef.requiredOption(`--${opt.name} <value>`, opt.description, noRepeatParser(opt.name));
        } else {
          commandDef.option(`--${opt.name} <value>`, opt.description, noRepeatParser(opt.name));
        }
      }
      const helpLines = [
        `\nGraph endpoint: ${cmd.meta.graphMethod} ${cmd.meta.graphPathTemplate}`,
        `Microsoft Learn: ${cmd.meta.graphDocsUrl}`,
        ...(cmd.meta.stability === 'experimental'
          ? [
              '\nStability: experimental — rides a Microsoft-internal substrate (chatsvcagg / IC3) that is not in the public Graph API and can break on a Teams web-client update. Prefer a stable sibling when one exists.',
            ]
          : []),
        ...(cmd.meta.pagination ? [`\nPagination: ${paginationHintFor(cmd.meta.paginationStrategy)}`] : []),
        ...(cmd.meta.bodyTemplate ? [`\nRequest body:\n  ${cmd.meta.bodyTemplate}`] : []),
        `\nExample:\n  ${cmd.meta.example}`,
      ];
      commandDef.addHelpText('after', helpLines.join('\n'));
      commandDef.action(async (opts: Record<string, string>) => {
        const globals = program.opts<{ outputPath?: string; outputDir?: string }>();
        const result = await runRegistryCommand({ graph, fs }, { name, command: cmd, params: opts, outputPath: globals.outputPath, outputDir: globals.outputDir, surface: 'cli' });
        if (result.ok) {
          renderOut(result.value, buildRenderContext(name, cmd, 'cli', opts));
          return;
        }
        fail(result.error.message, result.error.code, result.error.source, result.error.retryAfterSeconds);
      });
    }
  }

  return program;
};

export { buildCli };
export type { BuildCliDeps };
