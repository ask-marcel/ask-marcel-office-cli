import { describe, expect, it } from 'bun:test';
import type { GraphClient, GraphError } from '../infra/graph-client.ts';
import type { FileSystem } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { buildMediaSamples } from '../test-helpers/office-fixtures.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { commands } from '../use-cases/commands/index.ts';
import { buildCli } from './cli.ts';

const captureStream = async (stream: 'stdout' | 'stderr', run: () => void | Promise<unknown>): Promise<string> => {
  const target = process[stream];
  const original = target.write.bind(target);
  let captured = '';
  const swap = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  target.write = swap;
  try {
    await run();
  } finally {
    target.write = original;
  }
  return captured;
};

const okGraph = (value: unknown): GraphClient =>
  fakeGraphClient({
    get: async () => ({ ok: true, value }),
    post: async () => ({ ok: true, value }),
    getBinary: async () => ({ ok: true, value }),
    fetchUrl: async () => ({ ok: true, value }),
    put: async () => ({ ok: true, value }),
    delete: async () => ({ ok: true, value }),
  });

const errGraph = (error: GraphError): GraphClient =>
  fakeGraphClient({
    get: async () => ({ ok: false, error }),
    post: async () => ({ ok: false, error }),
    getBinary: async () => ({ ok: false, error }),
    fetchUrl: async () => ({ ok: false, error }),
    put: async () => ({ ok: false, error }),
    delete: async () => ({ ok: false, error }),
  });

describe('buildCli command surface', () => {
  it('top-level description derives endpoint counts from the registry, and never labels a search POST as a mutation (F-03)', () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const getCount = Object.values(commands).filter((c) => c.meta.graphMethod === 'GET' && c.meta.effect === 'read').length;
    const searchPosts = Object.entries(commands)
      .filter(([, c]) => c.meta.graphMethod === 'POST' && c.meta.effect === 'read')
      .map(([n]) => n);
    const description = cli.description();
    expect(description).toContain(`${getCount} GET endpoints`);
    // every search POST is listed in the "search POST" clause
    for (const name of searchPosts) expect(description).toContain(name);
    // the old bug: every POST was labelled "searches, not mutations", which
    // mislabelled create-mail-draft. That phrasing must be gone.
    expect(description).not.toContain('searches, not mutations');
  });

  it('renders a Graph error in text format with `error:` + `source: graph` (envelope-symmetry fix — round 2 — stamps the source even when the hint table did not match)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      graph: errGraph({ type: 'api_error', status: 404, message: 'not found' }),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    // The api_error → 'graph' mapping in `sourceFromGraphError` makes `source`
    // present on every Graph-originating failure, hint-matched or not. The
    // 404 message "not found" has no rule match → no `hint:` line → but the
    // `source: graph` line IS stamped from the explicit fallback.
    expect(out).toBe('error: not found\nsource: graph\n');
  });

  it('threads the Retry-After interval from a throttled Graph 429 into the JSON error envelope so a crawler can honor the backoff', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      graph: errGraph({ type: 'api_error', status: 429, message: 'TooManyRequests: Too many requests', code: 'TooManyRequests', retryAfterSeconds: 120 }),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-current-user']));
    const parsed = JSON.parse(out.trim()) as { ok: false; retryAfterSeconds?: number };
    expect(parsed.ok).toBe(false);
    expect(parsed.retryAfterSeconds).toBe(120);
  });

  it('runs a generic Graph command and renders the value as JSON', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({ displayName: 'Jordan' }), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    expect(out).toContain('Jordan');
  });

  // 2026-07-24: one name per flag. The alias spellings this block used to
  // exercise (--task-list-id, --id, --start/--end, --body-content) were removed.
  it('routes the canonical --todo-task-list-id to the task-list path', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
    });
    const logger = createLoggerFake();
    const cli = buildCli({ graph: captureGraph, logger, fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'list-todo-tasks', '--todo-task-list-id', 'AAMkABC']));
    expect(capturedPath).toBe('/me/todo/lists/AAMkABC/tasks');
  });

  it('convert-local-file-to-markdown reads the file through the CLI-wired filesystem and never touches Graph', async () => {
    let graphCalled = false;
    const tripwireGraph: GraphClient = fakeGraphClient({
      get: async () => {
        graphCalled = true;
        return { ok: true, value: {} };
      },
    });
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'name,age\nAlice,30');
    const cli = buildCli({ graph: tripwireGraph, logger: createLoggerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'convert-local-file-to-markdown', '--path', '/work/data.csv'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: { text: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.text).toContain('| Alice | 30 |');
    expect(graphCalled).toBe(false);
  });

  it('stamps a failed local-only command with source `cli`, not `graph` — it never touched Graph (F-02)', async () => {
    const fs = createFileSystemFake(); // nothing seeded → readBytes returns a not-found error
    const cli = buildCli({ graph: okGraph({}), logger: createLoggerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'convert-local-file-to-markdown', '--path', '/work/missing.csv'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: boolean; source?: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.source).toBe('cli');
  });

  it('routes the canonical --start-date-time/--end-date-time window to the calendarView path', async () => {
    let captured = '';
    const graph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        captured = path;
        return { ok: true, value: { value: [] } };
      },
    });
    const cli = buildCli({ graph, logger: createLoggerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', 'list-calendar-view', '--start-date-time', '2026-04-01T00:00:00Z', '--end-date-time', '2026-05-01T00:00:00Z'])
    );
    expect(captured).toContain('/me/calendarView?startDateTime=');
  });

  it('help-json rejects an explicit --output text (manifest is JSON by contract — audit )', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'text', 'help-json']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('help-json always emits JSON');
    expect(out).not.toContain('"commands"');
  });

  it('help-json still works when --output is left at its default text value (no explicit flag)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', 'help-json']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('"commands"');
  });

  // — compact default `--help` + help-json projections
  // (--terse / --category). These tests fix the discoverability gap an LLM
  // hits today: `--help` returns ~60 KB so the model is forced to truncate
  // or dump-to-disk before it can read the listing; `help-json` returns
  // 370 KB unfiltered. v1.4.0 surface-consolidation: the `--verbose` opt-out
  // was dropped — it was a one-trick toggle on this top-level listing only,
  // and `help-json --terse` covers the same need with a structured payload.
  //
  // Notes on test mechanics:
  // - Triggering Commander's `--help` flag directly would call `process.exit(0)`
  // even with `exitOverride` (the override returns, Commander then calls
  // `process.exit(exitCode)`), killing the test process. The custom `help`
  // subcommand goes through `program.outputHelp()` which renders without
  // exiting, so we use it for assertion. Production behaviour for `--help`
  // is exercised by the bare-args test at the bottom of this file.
  it('compact default help listing truncates each subcommand description to its first sentence', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    // get-mail-message has a multi-sentence summary post-§A; the compact
    // help should carry only the first sentence. The post-first-sentence
    // "ships a slim default" prose is the marker that we successfully cut.
    expect(out).toContain('Get a single Outlook message by ID.');
    expect(out).not.toContain('ships a slim default');
    // Footer must point at the discovery surfaces.
    expect(out).toContain('help-json [--terse] [--category mail]');
  });

  it('compact help listing stays under the 20 KB token-budget ceiling (byte-count regression guard — replaces the dropped --verbose opt-out)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const compact = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    // First-sentence truncation keeps the listing compact. The pre-compaction
    // full-summary form ran ~60 KB; if compaction silently regresses (e.g.
    // compactSummary cuts wrong), this guard fires. 45 KB ceiling: the
    // listing is ~37 KB as of 2.2.0 (the -to-markdown renames widened the
    // name|alias column commander pads every row to), leaving headroom for
    // new commands while still firing well before the ~60 KB regression form.
    expect(compact.length).toBeLessThan(45 * 1024);
  });

  // The ceiling guard above measured a moving target for two years. Commander
  // sizes its description column from `process.stdout.columns`, and the byte
  // count is not monotonic in width: on the 2.5.0 surface the same listing
  // rendered 35 KB at 80 columns, 64 KB at 100 and 46 KB at 140. Under `bun
  // test` in CI stdout is a pipe, so Commander fell back to 80 and the guard
  // was always green; run from an ordinary interactive terminal it failed, and
  // that is how a `npm publish` broke on 2026-09-06. A listing written to be
  // read by an LLM on a token budget must cost the same whoever renders it.
  it('renders the compact help listing at the same size whatever width the terminal reports', async () => {
    const logger = createLoggerFake();
    // Commander reads the width through `getOutHelpWidth`, which is
    // `process.stdout.isTTY ? process.stdout.columns : undefined`
    // (commander/lib/command.js). Under `bun test` stdout is a pipe, so both
    // have to be stubbed or this test measures the 80-column fallback four
    // times and passes against any bug.
    const originalColumns = process.stdout.columns;
    const originalIsTty = process.stdout.isTTY;
    const sizes: Array<number> = [];
    try {
      process.stdout.isTTY = true;
      for (const columns of [80, 100, 140, 200]) {
        process.stdout.columns = columns;
        const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
        sizes.push((await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']))).length);
      }
    } finally {
      process.stdout.columns = originalColumns;
      process.stdout.isTTY = originalIsTty;
    }
    expect(sizes.every((size) => size === sizes[0])).toBe(true);
  });

  it('--verbose is no longer a recognised top-level option (v1.4.0 surface-consolidation drop)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    let unknown = false;
    try {
      await cli.parseAsync(['node', 'ask-marcel-office', '--verbose', 'help']);
    } catch (e) {
      // Commander throws a CommanderError with code 'commander.unknownOption'
      // when it doesn't recognise a flag. exitOverride routes that through.
      unknown = (e as { code?: string }).code === 'commander.unknownOption';
    }
    expect(unknown).toBe(true);
  });

  it('help-json --terse strips per-command heavy fields (no options/example/graphPathTemplate/responseShape)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--terse']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"commands"');
    expect(out).not.toContain('"graphPathTemplate"');
    expect(out).not.toContain('"responseShape"');
    expect(out).not.toContain('"options"');
  });

  it('help-json --category mail filters the manifest down to the mail category', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--category', 'mail']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"get-mail-message"');
    // get-current-user is in the `user` category, must not leak in.
    expect(out).not.toContain('"get-current-user"');
  });

  it('help-json --terse --category mail composes: terse projection within a single category', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--terse', '--category', 'mail']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"get-mail-message"');
    expect(out).not.toContain('"options"');
    expect(out).not.toContain('"graphPathTemplate"');
  });

  it('help-json --category with an unknown name surfaces a structured `ok:false` envelope listing the valid categories', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--category', 'notarealcategory']);
      } catch {
        /* commander exits */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('Unknown --category "notarealcategory"');
    expect(parsed.error).toContain('mail');
    expect(parsed.error).toContain('drive');
  });

  it('stamps a machine-readable errorCode on an unknown --category rejection so an agent routes on the code, not the message', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--category', 'notarealcategory']);
      } catch {
        /* commander exits */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; errorCode?: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.errorCode).toBe('cli_unknown_category');
  });

  it('stamps a no_inlined_bytes errorCode when --output-path is used on a plain-JSON command with no body to write', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      graph: okGraph({ id: '1', displayName: 'Jordan Avery' }),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/out.json', 'get-current-user']));
    const parsed = JSON.parse(out.trim()) as { ok: false; errorCode?: string; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.errorCode).toBe('no_inlined_bytes');
    expect(parsed.error).toContain('did not return inlined bytes');
  });

  it('rejects a duplicate --output flag', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({ id: 'u1' }), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output', 'text', 'get-current-user']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('--output cannot be passed more than once');
  });

  it('rejects a single-value flag passed more than once', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({ value: [] }), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', 'list-folder-files', '--drive-id', 'd1', '--item-id', 'i1', '--filter', 'A', '--filter', 'B']);
      } catch {
        /* commander may exit on validation failure */
      }
    });
    expect(out).toContain('--filter cannot be passed more than once');
    expect(out).toContain('previous: "A"');
    expect(out).toContain('new: "B"');
  });

  it('still accepts the canonical flag name when the user does not use the alias', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
    });
    const logger = createLoggerFake();
    const cli = buildCli({ graph: captureGraph, logger, fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'list-todo-tasks', '--todo-task-list-id', 'AAMkXYZ']));
    expect(capturedPath).toBe('/me/todo/lists/AAMkXYZ/tasks');
  });

  it('renders the Graph error message when a generic Graph command fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      graph: errGraph({ type: 'api_error', status: 404, message: 'not found' }),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    expect(out).toContain('not found');
  });

  it('prints Markdown for a single command when the user runs `docs <cmd>`', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'get-current-user']));
    expect(out).toContain('# `get-current-user`');
    expect(out).toContain('## Example');
  });

  it('renders an unknown-command message when the user runs `docs` with a name that does not exist', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'this-is-not-a-real-command']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('this-is-not-a-real-command');
  });

  it('falls back to the unknown-command error when `docs help` is invoked (commander does not register `help` as a regular subcommand)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'help']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('help');
  });

  it("'help-json' subcommand prints the full machine-readable manifest (same shape as docs/commands.json) to stdout", async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, version: '1.0.0', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help-json']));
    const parsed = JSON.parse(out.trim()) as { package: string; version: string; commands: ReadonlyArray<{ name: string }> };
    expect(parsed.package).toBe('@ask-marcel/office-read');
    expect(parsed.version).toBe('1.0.0');
    expect(parsed.commands.length).toBeGreaterThan(100);
    expect(parsed.commands.some((c) => c.name === 'list-drives')).toBe(true);
  });

  // — `help-json` and `docs` used to bypass --output-path
  // entirely, neither writing the file nor surfacing a "not supported"
  // error. Both now honour the flag and write their text body. On error
  // (e.g. directory path) the same envelope as every other bytes-producing
  // command is surfaced (writeOrPrintText error path).
  it("'help-json' and 'docs' honour --output-path: write to disk + emit savedTo envelope; surface is_directory on a directory path", async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const make = (): ReturnType<typeof buildCli> => buildCli({ graph: okGraph({}), logger, version: '1.0.0', fs });
    const helpOut = await captureStream('stdout', () =>
      make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/manifest.json', 'help-json'])
    );
    const helpParsed = JSON.parse(helpOut.trim()) as { ok: true; data: { savedTo: string } };
    expect(helpParsed.data.savedTo).toBe('/work/test-output/manifest.json');
    const manifest = JSON.parse(fs.snapshot('/work/test-output/manifest.json') ?? '{}') as { commands: ReadonlyArray<unknown> };
    expect(manifest.commands.length).toBeGreaterThan(100);
    const docsOut = await captureStream('stdout', () =>
      make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/get-current-user.md', 'docs', 'get-current-user'])
    );
    const docsParsed = JSON.parse(docsOut.trim()) as { ok: true; data: { savedTo: string } };
    expect(docsParsed.data.savedTo).toBe('/work/test-output/get-current-user.md');
    expect(fs.snapshot('/work/test-output/get-current-user.md') ?? '').toContain('get-current-user');
    const failOut = await captureStream('stdout', async () => {
      try {
        await make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/', 'help-json']);
      } catch {
        /* expected */
      }
    });
    const failParsed = JSON.parse(failOut.trim()) as { ok: false; error: string };
    expect(failParsed.error).toContain('must be a file path, not a directory');
  });

  // — `--output-path` to a directory path used to surface
  // Node's `EISDIR: illegal operation on a directory`. Now it returns a
  // clear "must be a file path, not a directory" message.
  it('renders --output-path ending in / as "must be a file path, not a directory" rather than EISDIR', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync([
          'node',
          'ask-marcel-office',
          '--output',
          'json',
          '--output-path',
          '/work/test-output/',
          'download-drive-item-as-pdf',
          '--drive-id',
          'd1',
          '--item-id',
          'i1',
        ]);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('must be a file path, not a directory');
  });

  // — `*-as-pdf` commands silently falling back to source
  // bytes used to write a corrupt "PDF" to disk. Now the CLI rejects the
  // write with a clear message naming the actual content type.
  it('renders --output-path: passthrough source bytes (not converted PDF) — save as <contentType>, not .pdf, when the response carries passthrough:true', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const passthroughGraph: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'doc-v1.docx' } }),
      getBinaryElevated: async () => ({
        ok: true,
        value: { contentType: 'application/octet-stream', size: 12, base64: 'JVBERi0=' },
      }),
    };
    const cli = buildCli({ graph: passthroughGraph, logger, fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync([
          'node',
          'ask-marcel-office',
          '--output',
          'json',
          '--output-path',
          '/work/test-output/doc-v1.pdf',
          'download-drive-item-version',
          '--drive-id',
          'd1',
          '--item-id',
          'i1',
          '--version-id',
          '1.0',
          '--format',
          'pdf',
        ]);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('passthrough');
    expect(parsed.error).toContain('application/octet-stream');
    expect(fs.has('/work/test-output/doc-v1.pdf')).toBe(false);
  });

  // — when --output-path parent dir is missing /
  // non-writable, the raw Node `ENOENT: no such file or directory, mkdir
  // '/x'` used to leak through. Now translated to a clear message.
  it('translates a Node ENOENT/mkdir error from --output-path into a clear "parent directory missing or not writable" message', async () => {
    const logger = createLoggerFake();
    const fs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      chmod: async () => ({ ok: true as const, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed' as const, message: "ENOENT: no such file or directory, mkdir '/nonexistent-dir-no-write'" } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/nonexistent-dir-no-write/test.pdf',
        'download-drive-item-as-pdf',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('parent directory missing or not writable');
    expect(parsed.error).toContain('/nonexistent-dir-no-write');
  });

  it("renders a commander 'required option not specified' error for `search-mail-messages` without --query (asserts meta.options.required: true survives manifest → Commander wiring)", async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'search-mail-messages']);
      } catch {
        /* expected — commander throws via exitOverride on missing required option */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('required option');
    expect(parsed.error).toContain('--query');
  });

  it('routes commander parser errors (unknown option) to the JSON envelope on stdout, not stderr plain text (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const stderrOut = await captureStream('stderr', async () => {
      const stdoutOut = await captureStream('stdout', async () => {
        try {
          await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'list-drives', '--no-such-flag']);
        } catch {
          /* expected — commander throws on parser error after exitOverride */
        }
      });
      const parsed = JSON.parse(stdoutOut.trim()) as { ok: boolean; error: string };
      expect(parsed.ok).toBe(false);
      expect(parsed.error).toContain('unknown option');
      expect(parsed.error).toContain('--no-such-flag');
    });
    expect(stderrOut).toBe('');
  });

  it('routes commander parser errors as plain "error: ..." lines on stdout by default (text mode), nothing on stderr', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const stderrOut = await captureStream('stderr', async () => {
      const stdoutOut = await captureStream('stdout', async () => {
        try {
          await cli.parseAsync(['node', 'ask-marcel-office', 'list-drives', '--no-such-flag']);
        } catch {
          /* expected */
        }
      });
      expect(stdoutOut.startsWith('error: ')).toBe(true);
      expect(stdoutOut).toContain('--no-such-flag');
      expect(stdoutOut.endsWith('\n')).toBe(true);
    });
    expect(stderrOut).toBe('');
  });

  it('routes commander parser errors (unknown subcommand) to the JSON envelope on stdout (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'this-command-does-not-exist']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('this-command-does-not-exist');
  });

  it('routes commander parser errors (missing required option) to the JSON envelope on stdout (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    // get-mail-attachment keeps commander `requiredOption` (its flags carry no
    // aliases). get-mail-message can't be used here any more — its `--id` alias
    // moves required-ness to schema validation (commander no longer throws).
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-mail-attachment']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error.toLowerCase()).toContain('required');
  });

  it('rejects an invalid --output value with a plain "error: ..." line (text mode is the default for the error too)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'bogus', 'get-current-user']);
      } catch {
        /* expected */
      }
    });
    expect(out.startsWith('error: ')).toBe(true);
    expect(out).toContain("'bogus'");
  });

  it('global --output-path writes inline base64 bytes to disk and replaces base64 with savedTo in the envelope (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/work/test-output/may-deck.pdf',
        'download-drive-item-as-pdf',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { contentType: string; size: number; savedTo: string; base64?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/may-deck.pdf');
    expect(parsed.data.base64).toBeUndefined();
    const written = fs.snapshotBytes('/work/test-output/may-deck.pdf');
    expect(written).toBeDefined();
    if (written) expect(Array.from(written)).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('global --output-path writes a text body via writeText for markdown/plain-text returning commands (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const textGraph: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'notes.md' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'text/plain', size: 5, text: 'hello' } }),
    };
    const cli = buildCli({ graph: textGraph, logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/work/test-output/notes.md',
        'download-drive-item-as-markdown',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { savedTo: string; text?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/notes.md');
    expect(parsed.data.text).toBeUndefined();
    expect(fs.snapshot('/work/test-output/notes.md')).toBe('hello');
  });

  it('global --output-path is a no-op when the command returns plain JSON (no base64 / no text) — surfaces a clear error rather than silently no-op-ing (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const cli = buildCli({ graph: okGraph({ displayName: 'Jordan' }), logger, fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/profile.json', 'get-current-user']);
      } catch {
        /* commander may throw after exitOverride for explicit failures */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('--output-path');
    expect(parsed.error).toContain('did not return inlined bytes');
    expect(fs.has('/work/test-output/profile.json')).toBe(false);
  });

  it('global --output-path surfaces a write_failed error envelope when the filesystem rejects the write (e.g. permission denied) (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      chmod: async () => ({ ok: true as const, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed' as const, message: 'EACCES: permission denied, open' } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/root/forbidden.pdf', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('--output-path: write failed');
    expect(parsed.error).toContain('EACCES');
  });

  it('`ask-marcel-office` with NO subcommand prints --help to stdout instead of silently exiting 1', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office']));
    expect(out).toContain('Usage: ask-marcel-office-read');
    expect(out).toContain('help-json');
    expect(out).toContain('list-drives');
  });

  it('`help <unknown>` returns a JSON-envelope error rather than silently exiting (under --output json; )', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help', 'no-such-command']));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('no-such-command');
  });

  it('`help <known>` prints the markdown docs (alias of `docs <known>`)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help', 'list-drives']));
    expect(out).toContain('list-drives');
  });

  it('`help` with no argument prints the global --help text', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    expect(out).toContain('Usage: ask-marcel-office-read');
  });

  it('omitting --output-path leaves the JSON envelope unchanged (existing consumers still get base64) (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { base64?: string; savedTo?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.base64).toBe('JVBERi0=');
    expect(parsed.data.savedTo).toBeUndefined();
  });

  it('omitting --output-path on a binary command in text mode replaces base64 with a "use --output-path" hint so multi-MB blobs do not flood stdout', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 12345, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ graph: inlinedPdf, logger, fs });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1']));
    expect(out).toBe('binary: application/pdf, 12345 bytes — use --output-path to save\n');
  });

  const mediaGraph = async (): Promise<GraphClient> => {
    const bytes = await buildMediaSamples();
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'deck.pptx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/octet-stream', size: bytes.byteLength, base64: btoa(binary) } }),
    };
  };

  it('extract-drive-item-images --output-dir writes every image to the directory and replaces base64 with savedTo', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const cli = buildCli({ graph: await mediaGraph(), logger, fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', '/work/imgs'])
    );
    expect(out).toContain('savedTo');
    expect(out).toContain('/work/imgs/word_media_image1.png'); // full path flattened, not basename
    expect(out).not.toContain('base64');
    expect(Array.from(fs.snapshotBytes('/work/imgs/word_media_image1.png') ?? [])).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it('--output-dir on a command that returns no media array emits a clear error pointing at the image-extraction commands', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: okGraph({ displayName: 'Jordan' }), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-current-user', '--output-dir', '/work/imgs']));
    expect(out).toContain('did not return a media array');
    expect(out).toContain('extract-drive-item-images');
    expect(out).toContain('extract-mail-attachment-images');
  });

  it('rejects an empty --output-dir explicitly', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ graph: await mediaGraph(), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', ''])
    );
    expect(out).toContain('directory argument is empty');
  });

  it('surfaces a filesystem write failure from --output-dir', async () => {
    const logger = createLoggerFake();
    const failingFs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' } }),
      chmod: async () => ({ ok: true, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed', message: 'ENOSPC: no space left on device' } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const cli = buildCli({ graph: await mediaGraph(), logger, fs: failingFs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', '/work/imgs'])
    );
    expect(out).toContain('write failed');
    expect(out).toContain('ENOSPC');
  });
});
