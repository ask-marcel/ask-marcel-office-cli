import { Command, InvalidArgumentError, Option } from 'commander';
import type { CommanderError } from 'commander';
import { didYouMean } from '../domain/closest-names.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import type { ErrorSource } from '../presenter/error-hints.ts';
import type { OutputFormat } from '../presenter/output.ts';
import { render, renderError } from '../presenter/output.ts';
import type { RenderContext } from '../presenter/render-to-string.ts';
import { fillNames } from '../use-cases/commands/bin-names.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { buildManifest, buildTerseManifest, filterManifestByCategory, renderSingleCommand } from '../use-cases/commands/docs-builders.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, paginationHintFor } from '../use-cases/commands/docs-render.ts';
import { firstSentence } from '../use-cases/commands/first-sentence.ts';
import { persistIfRequested } from '../use-cases/commands/output-path.ts';
import { describeWrites } from '../use-cases/commands/command-effect.ts';
import { setDateZone } from '../use-cases/commands/date-zone.ts';
import { commandSuggestion, unknownCommandMessage } from '../use-cases/commands/unknown-command.ts';
import { isValidTimeZone } from '../domain/iso-datetime.ts';
import { resolveDateZone } from './date-zone.ts';
import { buildRenderContext, describeOutputPathError, runRegistryCommand } from './run-registry-command.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';

/**
 * How a lifecycle command a package adds prints: the same output and failure
 * paths as every command. `fail` fills the bin placeholders of its message;
 * `print` prints the data as it is, so a command fills its own (bin-names.ts).
 */
type CliOutput = {
  readonly print: (data: unknown) => void;
  readonly fail: (message: string, code?: string, source?: ErrorSource) => void;
};

/*
 * The CLI builder is core (package split, step 13): it builds the CLI of the
 * package it is given and imports no registry and no lifecycle command. The
 * composition root hands it the package (`cli.ts` for the single package).
 */
type CliShellDeps = {
  readonly commandPackage: CommandPackage;
  readonly graph: GraphClient;
  readonly logger: Logger;
  readonly fs: FileSystem;
  readonly version?: string;
  readonly onCommandError?: () => void;
  /** Adds the package's own lifecycle commands (`login`…) under the Lifecycle group, before `docs`. */
  readonly addLifecycleCommands?: (program: Command, output: CliOutput) => void;
};

const buildCliShell = (deps: CliShellDeps): Command => {
  const { commandPackage, graph, logger, fs, version } = deps;
  const { registry, binName } = commandPackage;
  const cliRemedy = `Run \`${binName} --help\` to list every command.`;
  const program = new Command();

  const getFormat = (): OutputFormat => {
    const raw = program.opts<{ output?: string }>().output;
    return raw === 'json' || raw === 'raw-json' ? raw : 'text';
  };
  // The CLI's own text (help, docs, hints, remedies) names commands through the
  // bin placeholders; this package's names fill them (bin-names.ts).
  const fill = (text: string): string => fillNames(text, commandPackage);
  const renderOut = (data: unknown, sizeHintContext?: RenderContext): void => render(data, logger, getFormat(), sizeHintContext, commandPackage);
  const fail = (message: string, code?: string, source?: ErrorSource, retryAfterSeconds?: number): void => {
    renderError(message, getFormat(), code, source, retryAfterSeconds, commandPackage);
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
    fail(describeOutputPathError(persisted.error, commandName, 'cli', commandPackage));
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
  // Commander's own suggestion reaches only three edits of the whole name; ours
  // also matches by words (`--folder-id` → `--mail-folder-id`), so theirs is off.
  // Set before any subcommand exists: a subcommand copies the setting when created.
  program.showSuggestionAfterError(false);
  const globalFlags = (): ReadonlyArray<string> => program.options.flatMap((o) => (o.long === undefined ? [] : [o.long]));
  // Commander reports a missing required flag before an unknown one, so a mistyped
  // required flag (`--folder` for `--mail-folder-id`) is only in what was typed.
  const typoedFlagHint = (optionFlags: ReadonlyArray<string>, typed: ReadonlyArray<string>): string => {
    const unknown = typed.find((arg) => arg.startsWith('--') && !optionFlags.includes(arg.split('=', 1)[0] ?? arg));
    return unknown === undefined ? '' : didYouMean(unknown, optionFlags, `\`${unknown}\` is not a flag of this command. Did you mean`);
  };
  const suggestionFor = (err: CommanderError, optionFlags: ReadonlyArray<string>, typed: ReadonlyArray<string>): string => {
    if (err.code === 'commander.missingMandatoryOptionValue') return typoedFlagHint(optionFlags, typed);
    const command = /^error: unknown command '([^']+)'/.exec(err.message)?.[1];
    if (command !== undefined)
      return commandSuggestion(
        command,
        program.commands.map((c) => c.name()),
        commandPackage,
        'cli'
      );
    const flag = /^error: unknown option '([^']+)'/.exec(err.message)?.[1];
    return flag === undefined ? '' : didYouMean(flag, optionFlags);
  };
  // `optionFlags`: the flags a mistyped one is compared with — the command's own
  // plus the global ones for a subcommand, the global ones at the top level.
  // `typed`: the command's leftover arguments, where a mistyped flag waits.
  const onCommanderError = (err: CommanderError, optionFlags: ReadonlyArray<string>, typed: ReadonlyArray<string> = []): void => {
    if (err.code === 'commander.helpDisplayed' || err.code === 'commander.help') return;
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
    const suggestion = suggestionFor(err, optionFlags, typed);
    fail(suggestion === '' ? stripped : `${stripped}.${suggestion}`, err.code, 'cli');
    throw err;
  };
  program.exitOverride((err) => onCommanderError(err, globalFlags()));

  // previous default `--help` ran ~60 KB because the
  // top-level subcommand listing rendered each command's full summary (often
  // 2-3 sentences). Default `--help` now truncates each subcommand description
  // to its first sentence in the top-level listing (~43 KB total). Per-command
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
  // GET reads only: the PDF conversions are declared GET, but they are writes.
  const getEndpointCount = Object.values(registry).filter((c) => c.meta.graphMethod === 'GET' && c.meta.effect === 'read').length;
  // 2026-06-15 (F-03), extended 2026-07-04 and at package-split step 11: everything
  // whose `effect` is not `read` is a write (the mail drafts, and the PDF
  // conversions that upload a temporary file); everything else is a read or a
  // search. Derive BOTH lists, and the wording of each write, from the effect so
  // the read-only narrative can never drift from the registry. The previous code
  // took every POST command and labelled it "— searches, not mutations", which
  // silently mislabelled create-mail-draft as a search the moment it shipped (and
  // omitted the PATCH update-mail-draft entirely).
  const writeCount = Object.values(registry).filter((c) => c.meta.effect !== 'read').length;
  const searchPostNames = Object.entries(registry)
    .filter(([, c]) => c.meta.graphMethod === 'POST' && c.meta.effect === 'read')
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));
  const surfaceDescription = `Microsoft Graph CLI. Read-mostly by design — the ONLY writes are ${writeCount} commands. ${describeWrites(registry)}. The CLI cannot create or modify calendar items (there is no send-mail / send-draft / create-event / upload-file command). ${getEndpointCount} GET endpoints + ${searchPostNames.length} search POST (${searchPostNames.join(', ')}). Safe default for LLM autonomy.`;

  program
    .name(binName)
    .description(surfaceDescription)
    // override the help-formatter's subcommand
    // description renderer to compact long summaries down to their first
    // sentence in the TOP-LEVEL `ask-marcel-office --help` listing. Per-subcommand
    // `ask-marcel-office <cmd> --help` is untouched — Commander only consults
    // `subcommandDescription` when formatting parent's child list.
    .configureHelp({
      // Pin the wrap width instead of inheriting `process.stdout.columns`.
      // Commander sizes the description column from the terminal, and the
      // resulting byte count is not monotonic in width: measured on the 2.5.0
      // surface, `help` renders 35 KB at 80 columns, 64 KB at 100 and 46 KB at
      // 140. This listing exists to be read by an LLM on a token budget, so its
      // size must be a property of the command surface, not of whoever's
      // terminal happens to be attached. It also made the byte-count guard in
      // cli.test.ts pass in CI (stdout is a pipe, so Commander falls back to 80)
      // and fail in any normal interactive terminal.
      helpWidth: 80,
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
        const ALLOWED: ReadonlyArray<string> = ['text', 'json', 'raw-json'];
        let outputSeen = false;
        return new Option(
          '--output <format>',
          'Output format. `text` (default, LLM-readable YAML-ish lines, ~30-60% fewer tokens on listings; errors render as `error: <message>`). `json` preserves the `{ok, data, nextLink?, deltaLink?, count?}` envelope for tool-chaining where unambiguous field extraction matters. `raw-json` prints the `data` payload alone, for piping into `jq` or a script: no envelope, no size hints, no paging cursors (use `json` to page); a failure still prints the JSON error envelope.'
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

  // Named days (`today`, `monday`, `start-of-month`) are wall-clock notions:
  // a user at UTC+8 asking for `today` at 02:00 means their day, not the one
  // UTC is still on. The zone comes from `--tz`, then `ASKMARCEL_TZ`, then the
  // machine, and is set once before any subcommand parses its dates.
  program.option(
    '--tz <zone>',
    "Globally available. The IANA time zone in which named days and boundaries resolve (`today`, `yesterday`, `monday`, `start-of-week`); instants such as `2026-04-01T10:00:00Z` and offsets such as `7d` are unaffected. Default: `ASKMARCEL_TZ`, else the machine's zone.",
    (value: string): string => {
      if (!isValidTimeZone(value)) throw new InvalidArgumentError('Not a known IANA time zone (e.g. Europe/Amsterdam, Asia/Shanghai, UTC).');
      return value;
    }
  );
  program.hook('preAction', () => {
    const resolved = resolveDateZone(program.opts<{ tz?: string }>().tz);
    setDateZone(resolved.zone);
    if (resolved.warning !== undefined) logger.warn('date_zone_fallback', { warning: resolved.warning });
  });

  // explicit pointers to per-command help and the
  // machine-readable manifest. Without this, an LLM that hits the compact
  // top-level help has no in-band signal that `help-json` even exists.
  program.addHelpText(
    'after',
    [
      '',
      `For full per-command help:   ${binName} <command> --help`,
      `For machine-readable docs:   ${binName} help-json [--terse] [--category mail]`,
      `For per-command Markdown:    ${binName} docs <command>`,
      `For the installed version:   ${binName} --version   (alone: on a command it is an unknown flag)`,
      // `mcp` is intercepted in main.ts, not registered as a subcommand (that
      // would pull the SDK into every command's module graph), so Commander
      // cannot list it by itself. Name it here or it is undiscoverable.
      `To serve this over MCP:      ${binName} mcp   (stdio; see README for \`claude mcp add\`)`,
    ].join('\n  ')
  );

  // bare `ask-marcel-office` (no subcommand) used to silently
  // exit 1 with zero output. We intercept that case BEFORE Commander parses
  // so we don't break the existing `unknown subcommand` error path. Hooked on
  // `preAction` of every subcommand would be wrong (it never fires for the
  // bare case); instead we override `parseAsync` itself.
  const originalParseAsync = program.parseAsync.bind(program);
  // `--version` or `-V`, alone or beside an output format (`--output json --version`).
  const isVersionRequest = (userArgs: ReadonlyArray<string>): boolean => {
    const rest = userArgs.filter((arg, i) => arg !== '--output' && userArgs[i - 1] !== '--output' && !arg.startsWith('--output='));
    return rest.length === 1 && (rest[0] === '--version' || rest[0] === '-V');
  };
  program.parseAsync = async (argv?: readonly string[], options?: { readonly from?: 'node' | 'electron' | 'user' }) => {
    const args = argv ?? process.argv;
    const from = options?.from ?? 'node';
    const userArgsStart = from === 'node' || from === 'electron' ? 2 : 0;
    if (args.length <= userArgsStart) {
      program.outputHelp();
      return program;
    }
    // Commander's version flag was global: `diff-drive-item-versions … --version 3.0`
    // (a mistyped --version-id) printed the version, dropped the command and
    // exited 0. It is answered here, as the whole request, and nowhere else.
    if (isVersionRequest(args.slice(userArgsStart))) {
      process.stdout.write(`${version ?? '0.0.0'}\n`);
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
      const result = renderSingleCommand(commandPackage, commandName);
      if (result.ok) {
        process.stdout.write(`${fill(result.value)}\n`);
        return;
      }
      // pass an explicit `cli_unknown_command`
      // code so error-hints can match it structurally — gives the same
      // envelope shape as Commander's own `commander.unknownCommand` path.
      fail(unknownCommandMessage(result.error.name, result.error.available, cliRemedy, commandPackage, 'cli'), 'cli_unknown_command');
    });

  program
    .command('help-json')
    .description(
      `Print the machine-readable command manifest as JSON. **Use \`--terse --category <name>\` for fresh-session discovery** — that combo is the actual token-friendly path (~8 KB for one category, vs ~575 KB unfiltered). The unflagged form is the *full* reference (every option / example / response shape per command) and is well over 10× the size of \`${binName} --help\`; reach for it only after \`--terse\` has narrowed the search. \`--terse\` alone projects to \`{name, summary, category}\` with each summary compacted to its first sentence (~41 KB across all categories). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.`
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
          `help-json always emits JSON (that's the contract — the manifest is the LLM-consumable serialized form of every command's meta). Drop \`--output text\` for this command. To browse the manifest as Markdown, use \`${binName} docs <command>\` instead.`
        );
        return;
      }
      // --terse projects to `{name, summary, category}`;
      // --category filters to a single category (validated against
      // CATEGORY_LABELS). Both flags compose.
      const fullOrTerse = opts.terse === true ? buildTerseManifest(commandPackage, version ?? '0.0.0') : buildManifest(commandPackage, version ?? '0.0.0');
      if (opts.category !== undefined) {
        const filtered = filterManifestByCategory(fullOrTerse, opts.category);
        if (!filtered.ok) {
          fail(`Unknown --category "${filtered.error.category}". Available categories: ${filtered.error.available.join(', ')}.`, 'cli_unknown_category');
          return;
        }
        await writeOrPrintText(fill(JSON.stringify(filtered.value)), 'application/json', 'help-json');
        return;
      }
      // The manifest is the CLI's own text only, so its JSON is filled whole.
      await writeOrPrintText(fill(JSON.stringify(fullOrTerse)), 'application/json', 'help-json');
    });

  program.commandsGroup('Lifecycle:');

  // The lifecycle commands that act (login, status, logout in the single package)
  // come from the package, so this builder knows no auth and no registry.
  deps.addLifecycleCommands?.(program, { print: (data) => renderOut(data), fail });

  const docsCmd = program
    .command('docs')
    .description(
      'Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`). Lifecycle commands (login/status/logout/docs/help-json) are also covered — they ship as manifest entries under category `lifecycle`.'
    )
    .argument('<command>', `Command name to show docs for (run \`${binName} --help\` to list every command).`)
    .action(async (commandName: string) => {
      const result = renderSingleCommand(commandPackage, commandName);
      if (!result.ok) {
        // structured `cli_unknown_command`
        // code so the envelope matches the `help <unknown>` and
        // `commander.unknownCommand` paths — single branch for LLM consumers.
        fail(unknownCommandMessage(result.error.name, result.error.available, cliRemedy, commandPackage, 'cli'), 'cli_unknown_command');
        return;
      }
      await writeOrPrintText(fill(result.value), 'text/markdown', 'docs');
    });
  docsCmd.addHelpText(
    'after',
    [
      '',
      `Example:       ${binName} docs list-mail-messages`,
      `Lifecycle:     \`${binName} docs login\` (or logout / docs) prints the same --help that command would, so you can introspect lifecycle commands the same way.`,
    ].join('\n  ')
  );

  for (const category of CATEGORY_ORDER) {
    const entries = Object.entries(registry).filter(([, c]) => c.meta.category === category);
    if (entries.length === 0) continue;
    program.commandsGroup(`${CATEGORY_LABELS[category]}:`);
    for (const [name, cmd] of entries) {
      const commandDef = program.command(name).description(fill(cmd.meta.summary));
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
          commandDef.requiredOption(`--${opt.name} <value>`, fill(opt.description), noRepeatParser(opt.name));
        } else {
          commandDef.option(`--${opt.name} <value>`, fill(opt.description), noRepeatParser(opt.name));
        }
      }
      commandDef.exitOverride((err) => onCommanderError(err, [...cmd.meta.options.map((o) => `--${o.name}`), ...globalFlags()], commandDef.args));
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
      commandDef.addHelpText('after', fill(helpLines.join('\n')));
      commandDef.action(async (opts: Record<string, string>) => {
        const globals = program.opts<{ outputPath?: string; outputDir?: string }>();
        const result = await runRegistryCommand(
          { graph, fs, commandPackage },
          { name, command: cmd, params: opts, outputPath: globals.outputPath, outputDir: globals.outputDir, surface: 'cli' }
        );
        if (result.ok) {
          renderOut(result.value, buildRenderContext(name, cmd, 'cli', opts));
          return;
        }
        fail(result.error.message, result.error.code, result.error.source, result.error.retryAfterSeconds);
      });
    }
  }

  // The lifecycle commands (help-json, login, docs…) suggest from their own flags too.
  for (const sub of program.commands) {
    if (sub.name() in registry) continue;
    const own = sub.options.flatMap((o) => (o.long === undefined ? [] : [o.long]));
    sub.exitOverride((err) => onCommanderError(err, [...own, ...globalFlags()], sub.args));
  }

  return program;
};

export { buildCliShell };
export type { CliOutput, CliShellDeps };
