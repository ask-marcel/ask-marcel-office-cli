import { describe, expect, it } from 'bun:test';
import { fakeCommandPackage, SIGN_IN_ENTRY } from '../../test-helpers/command-package-fake.ts';
import { buildManifest, buildTerseManifest, renderSingleCommand } from './docs-builders.ts';

// The docs builders are core: each package hands them its own registry, its own
// lifecycle entries and its own name. Nothing of the single package may leak in.
const at = (): Date => new Date('2026-10-14T08:00:00Z');
const WIDGET_NAMES = ['find-widgets', 'get-widget-file', 'get-widget-images', 'list-widgets', 'save-widget-draft', 'sign-in'];

describe('the manifest of a package', () => {
  it('names the package and lists its registry commands and its own lifecycle entries, sorted', () => {
    const manifest = buildManifest(fakeCommandPackage(), '1.2.3', at);
    expect(manifest.package).toBe('@widgets/cli');
    expect(manifest.version).toBe('1.2.3');
    expect(manifest.generatedAt).toBe('2026-10-14T08:00:00.000Z');
    expect(manifest.commands.map((c) => c.name)).toEqual(WIDGET_NAMES);
  });

  it('carries the lifecycle entry of the package as the package wrote it', () => {
    const manifest = buildManifest(fakeCommandPackage(), '1.2.3');
    expect(manifest.commands.find((c) => c.name === 'sign-in')).toEqual(SIGN_IN_ENTRY);
  });

  it('holds only the registry when the package has no lifecycle entry', () => {
    const manifest = buildManifest(fakeCommandPackage({ lifecycle: [] }), '1.2.3');
    expect(manifest.commands.map((c) => c.name)).toEqual(WIDGET_NAMES.filter((name) => name !== 'sign-in'));
  });

  it('gives the terse view the same package name and the same entries', () => {
    const terse = buildTerseManifest(fakeCommandPackage(), '1.2.3', at);
    expect(terse.package).toBe('@widgets/cli');
    expect(terse.generatedAt).toBe('2026-10-14T08:00:00.000Z');
    expect(terse.commands.map((c) => c.name)).toEqual(WIDGET_NAMES);
    expect(terse.commands.find((c) => c.name === 'sign-in')).toEqual({ name: 'sign-in', summary: 'Sign in to the widget service.', category: 'lifecycle' });
  });
});

describe('the docs page of one command of a package', () => {
  it('renders a lifecycle entry of the package', () => {
    const page = renderSingleCommand(fakeCommandPackage(), 'sign-in');
    expect(page.ok && page.value).toContain('Sign in to the widget service.');
  });

  it('renders a registry command of the package', () => {
    const page = renderSingleCommand(fakeCommandPackage(), 'get-widget-file');
    expect(page.ok && page.value).toContain('Download a widget file.');
  });

  it('refuses a name the package does not ship, and offers only what the package ships', () => {
    const page = renderSingleCommand(fakeCommandPackage(), 'login');
    expect(page).toEqual({ ok: false, error: { type: 'unknown_command', name: 'login', available: WIDGET_NAMES } });
  });
});
