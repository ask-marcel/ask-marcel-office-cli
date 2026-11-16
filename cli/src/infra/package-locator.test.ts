import { describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import { createPackageLocator } from './package-locator.ts';

/*
 * The real locator asks the runtime where a package is, from the folder of the
 * facade. The write package is a dependency of this folder, so the workspace
 * has it.
 */

describe('the locator of an installed package', () => {
  it('finds the package.json of a package that this folder depends on', async () => {
    const found = await createPackageLocator().manifestPath('@ask-marcel/office-write');
    if (!found.ok) throw new Error('the write package was not found');
    expect(basename(found.value)).toBe('package.json');
    expect(JSON.parse(readFileSync(found.value, 'utf8'))).toMatchObject({ name: '@ask-marcel/office-write' });
  });

  it('gives not_installed for a package that is not installed', async () => {
    const found = await createPackageLocator().manifestPath('@ask-marcel/office-absent');
    expect(found).toEqual({ ok: false, error: { type: 'not_installed' } });
  });
});
