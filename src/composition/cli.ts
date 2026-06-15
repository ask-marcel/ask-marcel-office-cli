import { Command, InvalidArgumentError, Option } from 'commander';
import type { AuthManager } from '../infra/auth.ts';
import { createAuthManager } from '../infra/auth.ts';
import type { GraphClient, GraphError } from '../infra/graph-client.ts';
import type { ErrorSource } from '../presenter/error-hints.ts';
import type { OutputFormat } from '../presenter/output.ts';
import { render, renderError } from '../presenter/output.ts';
import { buildManifest, buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, paginationHintFor } from '../use-cases/commands/docs-render.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import * as login from '../use-cases/commands/login.ts';
import * as logout from '../use-cases/commands/logout.ts';
import type { OutputDirError, OutputPathError } from '../use-cases/commands/output-path.ts';
import { persistIfRequested, persistMediaIfRequested } from '../use-cases/commands/output-path.ts';
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

  const getFormat = (): OutputFormat => {
    const raw = program.opts<{ output?: string }>().output;
    return raw === 'json' ? 'json' : 'text';
  };
  const renderOut = (data: unknown): void => render(data, logger, getFormat());
  const fail = (message: string, code?: string, source?: ErrorSource): void => {
    renderError(message, getFormat(), code, source);
    deps.onCommandError?.();
  };

  // v1.4.0 fresh-pass #5 (round 2): map the discriminated GraphError type to
  // the presenter's ErrorSource. `api_error` and `network_error` both surface
  // Graph endpoint failures (the latter is "couldn't even reach Graph"), so
  // both classify as `graph`. `validation_error` is Zod / use-case-side
  // schema rejection. `auth_failed` is a CLI-side concern (the remedy is
  // `ask-marcel login`, not a Graph operation). When the hint table matches
  // a more specific source (e.g. `substrate` for chatsvcagg-coded errors),
  // the hint's source wins — this explicit source is just the fallback so
  // bare envelopes still carry the classifier.
  const sourceFromGraphError = (error: GraphError): ErrorSource => {
    if (error.type === 'validation_error') return 'validation';
    if (error.type === 'auth_failed') return 'cli';
    return 'graph';
  };

  // Audit round-8 Wave E2: manifest-driven --output-path-supporting list.
  const bytesProducingCommands = Object.entries(cmdRegistry)
    .filter(([, c]) => c.meta.producesBytes === true)
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));

  // Parallel manifest-driven list for the --output-dir flag.
  const mediaProducingCommands = Object.entries(cmdRegistry)
    .filter(([, c]) => c.meta.producesMedia === true)
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));

  const formatOutputPathError = (error: OutputPathError, commandName: string): string => {
    if (error.type === 'no_inlined_bytes')
      return `--output-path: ${commandName} did not return inlined bytes — this flag works only with commands that produce a body to write. Supported: ${bytesProducingCommands.join(', ')}. Plain JSON commands (list-*, get-*-user, get-organization, etc.) don't have a body to write — drop the flag and use a shell redirect instead: \`ask-marcel ${commandName} ... > out.json\`.`;
    if (error.type === 'empty_path') return '--output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)';
    // Audit v1.0.0 §B11: paths ending in `/` or `\` look like a directory; reject upfront instead of Node's `EISDIR`.
    if (error.type === 'is_directory') return '--output-path: must be a file path, not a directory (paths ending in `/` or `\\` look like a directory).';
    // Audit v1.0.0 §B4: *-as-pdf fallbacks return source bytes with `passthrough:true`; refuse `.pdf` to avoid a corrupt save.
    if (error.type === 'passthrough_extension_mismatch')
      return `--output-path: response is passthrough source bytes (contentType: \`${error.contentType}\`), NOT a converted PDF. Save with the source extension matching that contentType, not \`${error.requestedExtension}\` — see the response's \`note\` field.`;
    // Audit v1.0.0 §B10: humanise the ENOENT/mkdir shape; preserve EACCES/ENOSPC verbatim — they're already actionable.
    const enoent = /^ENOENT:.*'([^']+)'/.exec(error.message);
    if (enoent !== null) return `--output-path: parent directory missing or not writable: ${enoent[1]}`;
    return `--output-path: write failed: ${error.message}`;
  };

  const formatOutputDirError = (error: OutputDirError, commandName: string): string => {
    if (error.type === 'no_media')
      return `--output-dir: ${commandName} did not return a media array — this flag works only with the image-extraction commands. Supported: ${mediaProducingCommands.join(', ')}.`;
    if (error.type === 'empty_path') return '--output-dir: directory argument is empty (likely a shell-quoting mistake — pass a real directory path)';
    return `--output-dir: write failed: ${error.message}`;
  };

  // Audit v1.0.0 §B5: route help-json / docs through persistIfRequested so --output-path is honoured (was silently ignored).
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
    fail(formatOutputPathError(persisted.error, commandName));
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
    // Audit Jane-session §2 follow-up: pass Commander's structured code
    // (`commander.unknownOption`, `commander.missingMandatoryOptionValue`,
    // etc.) through to the error renderer so the hint table can match it and
    // surface `hint:` / `source:` for the CLI-input failure cases too. Prior
    // behaviour dropped the code, leaving Commander errors as bare
    // `{ok:false,error}` envelopes — inconsistent with the Graph error shape.
    // v1.4.0 fresh-pass #5 (round 2): also stamp `source: 'cli'` explicitly
    // so the envelope shape stays stable even if a future Commander error
    // code variant doesn't have a matching rule yet.
    fail(stripped, err.code, 'cli');
    throw err;
  });

  // Audit Jane-session §B: previous default `--help` ran ~60 KB because the
  // top-level subcommand listing rendered each command's full summary (often
  // 2-3 sentences). Default `--help` now truncates each subcommand description
  // to its first sentence in the top-level listing (~6 KB total). Per-command
  // `ask-marcel <cmd> --help` is never compacted — it always shows the full
  // description plus the `addHelpText` block, and `help-json` ships the full
  // summary unchanged. v1.4.0 surface-consolidation: the `--verbose` opt-out
  // was dropped (it was a one-trick toggle on this top-level listing only;
  // `help-json --terse` covers the same need with a structured payload).

  // Compact a subcommand's description to its first sentence (everything up
  // to the first period followed by a space, or the first newline). Returns
  // the full text untouched when no sensible cut-point exists, so we never
  // produce a less-readable result than the original. The cut keeps the
  // trailing period for natural reading.
  const compactSummary = (full: string): string => {
    const newlineIdx = full.indexOf('\n');
    const periodIdx = full.search(/\. /);
    const candidates = [newlineIdx, periodIdx].filter((i) => i > 0);
    if (candidates.length === 0) return full;
    const cut = Math.min(...candidates);
    // Keep the period itself when cutting on `. ` so the line ends naturally;
    // strip a stray period for the newline case.
    return full.charAt(cut) === '.' ? full.slice(0, cut + 1) : full.slice(0, cut);
  };

  // QA fix: derive the advertised endpoint counts from the registry so the
  // top-level description can never drift from the manifest again. The
  // hardcoded "164 GET + 1 POST" literal had gone stale (real: 169 GET + 2
  // POST) when search-all-accessible-sites was added without updating it.
  const getEndpointCount = Object.values(cmdRegistry).filter((c) => c.meta.graphMethod === 'GET').length;
  const postCommandNames = Object.entries(cmdRegistry)
    .filter(([, c]) => c.meta.graphMethod === 'POST')
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));
  const surfaceDescription = `Microsoft Graph CLI. READ-ONLY: this CLI cannot send mail, modify or create calendar items, write files, or perform any mutating action — there is no send-mail / create-event / upload-file command. ${getEndpointCount} GET endpoints + ${postCommandNames.length} POST (${postCommandNames.join(', ')} — searches, not mutations). Safe default for LLM autonomy.`;

  program
    .name('ask-marcel')
    .description(surfaceDescription)
    .version(version ?? '0.0.0')
    // Audit Jane-session §B: override the help-formatter's subcommand
    // description renderer to compact long summaries down to their first
    // sentence in the TOP-LEVEL `ask-marcel --help` listing. Per-subcommand
    // `ask-marcel <cmd> --help` is untouched — Commander only consults
    // `subcommandDescription` when formatting parent's child list.
    .configureHelp({
      subcommandDescription: (cmd) => compactSummary(cmd.description()),
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
        // Audit round-8 Wave G: reject duplicate `--output` flags AND
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

  // Audit Jane-session §B: explicit pointers to per-command help and the
  // machine-readable manifest. Without this, an LLM that hits the compact
  // top-level help has no in-band signal that `help-json` even exists.
  program.addHelpText(
    'after',
    [
      '',
      'For full per-command help:   ask-marcel <command> --help',
      'For machine-readable docs:   ask-marcel help-json [--terse] [--category mail]',
      'For per-command Markdown:    ask-marcel docs <command>',
    ].join('\n  ')
  );

  // Audit v1.0.0 §2.3: bare `ask-marcel` (no subcommand) used to silently
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
      // Audit Jane-session §6 follow-up: pass an explicit `cli_unknown_command`
      // code so error-hints can match it structurally — gives the same
      // envelope shape as Commander's own `commander.unknownCommand` path.
      fail(`Unknown command "${result.error.name}". Run \`ask-marcel --help\` to list every command.`, 'cli_unknown_command');
    });

  program
    .command('help-json')
    .description(
      'Print the machine-readable command manifest as JSON. **Use `--terse --category <name>` for fresh-session discovery** — that combo is the actual token-friendly path (~12 KB for one category, vs ~370 KB unfiltered). The unflagged form is the *full* reference (every option / example / response shape per command) and is roughly 13× the size of `ask-marcel --help`; reach for it only after `--terse` has narrowed the search. `--terse` alone projects to `{name, summary, category}` (~62 KB across all categories). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.'
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
      // Audit round-8 Wave G1: --output text was silently honored on
      // help-json, contradicting the global flag's contract. The manifest IS
      // JSON; serializing as text has no use case. Reject only when the user
      // EXPLICITLY passed `--output text` (defaulting to text and getting
      // JSON anyway is fine — that's the historical behavior).
      const outputSource = program.getOptionValueSource('output');
      if (outputSource === 'cli' && getFormat() === 'text') {
        fail(
          "help-json always emits JSON (that's the contract — the manifest is the LLM-consumable serialized form of every command's meta). Drop `--output text` for this command. To browse the manifest as Markdown, use `ask-marcel docs <command>` instead."
        );
        return;
      }
      // Audit Jane-session §B: --terse projects to `{name, summary, category}`;
      // --category filters to a single category (validated against
      // CATEGORY_LABELS). Both flags compose.
      const fullOrTerse =
        opts.terse === true
          ? buildTerseManifest(cmdRegistry, 'ask-marcel-office-cli', version ?? '0.0.0')
          : buildManifest(cmdRegistry, 'ask-marcel-office-cli', version ?? '0.0.0');
      if (opts.category !== undefined) {
        const filtered = filterManifestByCategory(fullOrTerse, opts.category);
        if (!filtered.ok) {
          fail(`Unknown --category "${filtered.error.category}". Available categories: ${filtered.error.available.join(', ')}.`);
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
    .description('Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback).')
    .option('--use-extension', 'Use browser extension for authentication (requires Ask Marcel Companion extension)')
    .action(async (opts: { useExtension?: boolean }) => {
      // Default: use Playwright (skipSystemBrowser: true)
      // --use-extension: use browser extension (skipSystemBrowser: false, usePlaywrightFallback: false)
      // Cross-platform home directory
      const homeDir = process.env.HOME || process.env.USERPROFILE || '';
      const useExtension = opts.useExtension ?? false;
      const loginAuth = createAuthManager({
        cachePath: homeDir ? `${homeDir}/.ask-marcel/token-cache.json` : '',
        logger,
        fs,
        skipSystemBrowser: !useExtension,
        usePlaywrightFallback: !useExtension,
      });
      const result = await login.execute(loginAuth);
      if (!result.ok) {
        fail(result.error.type === 'auth_cancelled' ? 'Authentication cancelled' : result.error.message);
        return;
      }
      // Login-fix round-1 Wave D: surface the elevated-capture outcome so
      // an LLM consumer can predict whether the elevated-dependent
      // commands (list-chats / get-chat / list-chat-members /
      // historical-version downloads) will work without invoking them.
      // The outcome is null when getAccessToken hit cache (no browser
      // step ran in this process) — leave the elevated field unset in
      // that case so consumers don't think the elevated state was
      // tested.
      const outcome = auth.getLastElevatedOutcome();
      const envelope: { status: 'authenticated'; elevated?: 'captured' | 'failed'; elevatedReason?: string } = { status: 'authenticated' };
      if (outcome !== null) {
        if (outcome.captured) {
          envelope.elevated = 'captured';
        } else {
          envelope.elevated = 'failed';
          envelope.elevatedReason = outcome.reason;
        }
      }
      renderOut(envelope);
    });
  loginCmd.addHelpText(
    'after',
    [
      '',
      'Example:       ask-marcel login',
      '               ask-marcel login --use-playwright  (fallback to Playwright if extension fails)',
      'Token cache:   ~/.ask-marcel/token-cache.json (access + refresh tokens, JSON, 0600).',
      'Browser ext:   Install "Ask Marcel Companion" for faster login (no Playwright needed).',
      '               See browser-extension/README.md for installation instructions.',
      'Browser data:  ~/.ask-marcel/browser-profile/ (Playwright persistent context, only with --use-playwright).',
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
      if (result.ok) renderOut({ status: 'logged_out' });
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
      if (result.ok) renderOut({ status: 'updated', via: manager });
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
    .action(async (commandName: string) => {
      const result = renderSingleCommand(cmdRegistry, commandName);
      if (!result.ok) {
        // Audit Jane-session §6 follow-up: structured `cli_unknown_command`
        // code so the envelope matches the `help <unknown>` and
        // `commander.unknownCommand` paths — single branch for LLM consumers.
        fail(`Unknown command "${result.error.name}". Run \`ask-marcel --help\` to list every command.`, 'cli_unknown_command');
        return;
      }
      await writeOrPrintText(result.value, 'text/markdown', 'docs');
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
      for (const opt of cmd.meta.options) {
        if (opt.aliases && opt.aliases.length > 0) {
          // When aliases exist we can't use `requiredOption` for the canonical
          // — Commander would reject alias-only invocations (the canonical
          // long flag would be missing). Schema validation
          // (z.string().min(1)) enforces required-ness instead.
          commandDef.option(`--${opt.name} <value>`, opt.description, noRepeatParser(opt.name));
          for (const alias of opt.aliases) {
            commandDef.option(`--${alias.name} <value>`, `(alias for --${opt.name})`, noRepeatParser(alias.name));
          }
        } else if (opt.required) {
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
        const normalized: Record<string, string> = { ...opts };
        // Audit round-7 B4: track which alias the user actually typed so the
        // post-validation error message references the flag they typed rather
        // than the canonical schema name (e.g. `--query is empty` not
        // `--title-substring is empty`).
        const aliasUsedFor: Record<string, string> = {};
        for (const opt of cmd.meta.options) {
          for (const alias of opt.aliases ?? []) {
            const aliasValue = opts[alias.key];
            if (typeof aliasValue === 'string' && !Object.hasOwn(opts, opt.key)) {
              normalized[opt.key] = aliasValue;
              aliasUsedFor[opt.name] = alias.name;
            }
          }
        }
        // `convert-local-file` is the one command whose input is the local
        // filesystem, not Graph — route it to executeLocal with the
        // composition-selected FileSystem (the same instance --output-path uses).
        const result = cmd.executeLocal !== undefined ? await cmd.executeLocal(fs, normalized) : await cmd.execute(graph, normalized);
        if (!result.ok) {
          let message = result.error.message;
          for (const [canonical, alias] of Object.entries(aliasUsedFor)) {
            message = message.replaceAll(`--${canonical}`, `--${alias}`);
          }
          fail(message, result.error.code, sourceFromGraphError(result.error));
          return;
        }
        const outputDir = program.opts<{ outputDir?: string }>().outputDir;
        if (outputDir !== undefined) {
          const persistedMedia = await persistMediaIfRequested(fs, outputDir, result.value);
          if (persistedMedia.ok) {
            renderOut(persistedMedia.value);
            return;
          }
          fail(formatOutputDirError(persistedMedia.error, name));
          return;
        }
        const outputPath = program.opts<{ outputPath?: string }>().outputPath;
        const persisted = await persistIfRequested(fs, outputPath, result.value);
        if (persisted.ok) {
          renderOut(persisted.value);
          return;
        }
        fail(formatOutputPathError(persisted.error, name));
      });
    }
  }

  return program;
};

export { buildCli };
export type { BuildCliDeps };
