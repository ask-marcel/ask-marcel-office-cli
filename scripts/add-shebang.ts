#!/usr/bin/env bun
/*
 * Build post-step: prepend `#!/usr/bin/env node` to dist/cli.js and chmod 0o755
 * so that npm symlinks it as an executable global bin. dist/token.js, the token
 * helper's own entry, gets the same treatment so a caller can name it directly
 * (`ASKMARCEL_TOKEN_COMMAND=/path/to/dist/token.js`).
 *
 * Bun build does not insert shebangs, so this runs after `bun build`.
 */

import { chmodSync } from 'node:fs';

const TARGETS = ['dist/cli.js', 'dist/token.js'];
const SHEBANG = '#!/usr/bin/env node\n';

for (const target of TARGETS) {
  const file = Bun.file(target);
  if (!(await file.exists())) {
    process.stderr.write(`add-shebang: ${target} not found — did bun build run?\n`);
    process.exit(1);
  }

  const existing = await file.text();
  if (existing.startsWith('#!')) {
    process.stderr.write(`add-shebang: ${target} already has a shebang — leaving as-is\n`);
  } else {
    await Bun.write(target, SHEBANG + existing);
    process.stderr.write(`add-shebang: shebang prepended to ${target}\n`);
  }

  chmodSync(target, 0o755);
  process.stderr.write(`add-shebang: ${target} marked executable\n`);
}
