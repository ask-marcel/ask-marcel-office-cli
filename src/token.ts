import { fileURLToPath } from 'node:url';
import { version } from '../package.json' with { type: 'json' };
import { runTokenHelper } from './composition/token-helper.ts';

/*
 * The `token` helper's own entry, built to dist/token.js: read and write spawn it
 * for every bearer they do not hold, so it loads only the auth ladder, the file
 * system and the token endpoint call, never the CLI (commander, winston, the
 * update notifier, the output renderer). It records itself as the helper on
 * every run. `ask-marcel-office token ...` reaches the same code through main.ts.
 */
try {
  process.exitCode = await runTokenHelper({
    argv: process.argv.slice(2),
    location: { execPath: process.execPath, entry: fileURLToPath(import.meta.url), version },
  });
} catch (e) {
  // Only the error's kind: a message can quote what was being parsed, and a
  // token reaches no stream but the one stdout line.
  process.stderr.write(`[crash] the token helper stopped on an unexpected ${e instanceof Error ? e.name : 'throw'}; its message is withheld.\n`);
  process.exitCode = 1;
}
