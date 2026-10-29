import { describe, expect, it } from 'bun:test';
import { COMMAND_DIRECTORY, commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { commands } from './write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from './write-lifecycle-entries.ts';

/*
 * The meta invariants of meta.test.ts, held on the registry of the write package
 * (package split, phase 5, E6). Only what is true of a registry of writes is
 * here: the read-only pins (the substrate and elevated sets, the commands that
 * declare no scope, the runtime query flags) stay in meta.test.ts. No write
 * command has a runtime query flag, so every option of a write command is a
 * placeholder of its path or body template.
 */

const camelToKebab = (s: string): string => s.replaceAll(/([A-Z])/g, '-$1').toLowerCase();

const schemaKeys = (schema: Command['schema']): string[] => {
  const shape = (schema as unknown as { shape?: Record<string, unknown> }).shape;
  return shape ? Object.keys(shape) : [];
};

const placeholders = (template: string): string[] => Array.from(template.matchAll(/\{([a-z][a-z0-9-]*)\}/g), (m) => m[1] ?? '');

// True when a Zod field is `.optional()`, so meta.required stays in step with the schema.
const isOptionalSchemaField = (schema: Command['schema'], key: string): boolean => {
  const shape = (schema as unknown as { shape?: Record<string, { isOptional?: () => boolean }> }).shape;
  const field = shape?.[key];
  return typeof field?.isOptional === 'function' ? field.isOptional() : false;
};

const sortedNames = (matches: (command: (typeof commands)[string]) => boolean): ReadonlyArray<string> =>
  Object.entries(commands)
    .filter(([, command]) => matches(command))
    .map(([name]) => name)
    .toSorted((a, b) => a.localeCompare(b));

describe('command meta: invariants on every command of the write package', () => {
  it('classifies EXACTLY the four mail-draft commands as `draft` writes and the three PDF converters as `transient-upload` writes, and holds no `read` (F-03)', () => {
    expect(sortedNames((command) => command.meta.effect === 'draft')).toEqual(['create-forward-draft', 'create-mail-draft', 'create-reply-draft', 'update-mail-draft']);
    expect(sortedNames((command) => command.meta.effect === 'transient-upload')).toEqual([
      'convert-calendar-event-attachment-to-pdf',
      'convert-group-post-attachment-to-pdf',
      'convert-mail-attachment-to-pdf',
    ]);
    expect(sortedNames((command) => command.meta.effect === 'read')).toEqual([]);
  });

  it('flags no write command with a substrate or an elevated token: a write runs on the basic tier only', () => {
    expect(sortedNames((command) => command.meta.needsSubstrateToken !== undefined || command.meta.needsElevatedToken !== undefined)).toEqual([]);
  });

  it('declares its Graph scopes inline on every write command, so `help-json` can predict a 403 before the call', () => {
    expect(sortedNames((command) => (command.meta.scopesRequired ?? []).length === 0)).toEqual([]);
  });

  it('writes every scope in the PascalCase.Verb form Microsoft documents (e.g. `Mail.ReadWrite`, `Files.ReadWrite`)', () => {
    const scopePattern = /^[A-Z][A-Za-z]+(\.[A-Z][A-Za-z]+)*$/;
    const malformed = Object.entries(commands).flatMap(([name, command]) =>
      (command.meta.scopesRequired ?? []).filter((scope) => !scopePattern.test(scope)).map((scope) => `${name}: ${scope}`)
    );
    expect(malformed).toEqual([]);
  });

  for (const [name, command] of Object.entries(commands)) {
    describe(`meta for \`${name}\``, () => {
      it('has a non-empty summary', () => {
        expect(command.meta.summary.trim().length).toBeGreaterThan(0);
      });

      it('declares one CommandOptionMeta entry per Zod schema field', () => {
        const keys = schemaKeys(command.schema).toSorted((a, b) => a.localeCompare(b));
        const optionKeys = command.meta.options.map((o) => o.key).toSorted((a, b) => a.localeCompare(b));
        expect(optionKeys).toEqual(keys);
      });

      it('uses kebab-case `name` matching the camelCase `key` in every option, with a description', () => {
        for (const opt of command.meta.options) {
          expect(opt.name).toBe(camelToKebab(opt.key));
          expect(opt.description.trim().length).toBeGreaterThan(0);
        }
      });

      it('declares no aliases anywhere: one name per flag, one name per command (2026-07-24)', () => {
        for (const opt of command.meta.options) expect(Object.hasOwn(opt, 'aliases')).toBe(false);
        expect(Object.hasOwn(command.meta, 'commandAliases')).toBe(false);
      });

      it('references each option at least once across graphPathTemplate + bodyTemplate, and references nothing else', () => {
        const expected = Array.from(new Set(command.meta.options.map((o) => o.name))).toSorted((a, b) => a.localeCompare(b));
        const combined = `${command.meta.graphPathTemplate} ${command.meta.bodyTemplate ?? ''}`;
        const found = Array.from(new Set(placeholders(combined))).toSorted((a, b) => a.localeCompare(b));
        expect(found).toEqual(expected);
      });

      it('publishes a Microsoft Learn URL for the underlying Graph endpoint or guide', () => {
        expect(command.meta.graphDocsUrl).toMatch(/^https:\/\/learn\.microsoft\.com\/en-us\/graph\//);
      });

      it('provides a runnable example beginning with the write bin and its own name', () => {
        expect(commandRole(name)).toBe('write');
        expect(command.meta.example.startsWith(`{write-bin} ${name}`)).toBe(true);
      });

      it('has a non-empty category, graphMethod from the HTTP-verb set, non-empty graphPathTemplate, and non-empty responseShape when declared', () => {
        expect(command.meta.category.trim().length).toBeGreaterThan(0);
        expect(['GET', 'POST', 'PUT', 'DELETE', 'PATCH']).toContain(command.meta.graphMethod);
        expect(command.meta.graphPathTemplate.trim().length).toBeGreaterThan(0);
        if (command.meta.responseShape !== undefined) expect(command.meta.responseShape.trim().length).toBeGreaterThan(0);
      });

      it('pairs each option.required with the schema field optionality: required:true iff the Zod field is NOT .optional()', () => {
        for (const opt of command.meta.options) {
          const isOpt = isOptionalSchemaField(command.schema, opt.key);
          expect({ key: opt.key, required: opt.required }).toEqual({ key: opt.key, required: !isOpt });
        }
      });

      it('declares non-empty magicValue argumentHint values and only-`true` capability flags', () => {
        for (const opt of command.meta.options) {
          if (opt.argumentHint !== undefined) expect(opt.argumentHint.kind.trim().length).toBeGreaterThan(0);
          if (opt.argumentHint?.kind === 'magicValue') {
            expect(opt.argumentHint.values.length).toBeGreaterThan(0);
            for (const value of opt.argumentHint.values) expect(value.trim().length).toBeGreaterThan(0);
          }
        }
        for (const flag of [command.meta.producesBytes, command.meta.producesMedia, command.meta.pagination]) {
          if (flag !== undefined) expect(flag).toBe(true);
        }
      });

      it('summary references only flags that are registered as options on this command', () => {
        const flagsInSummary = Array.from(command.meta.summary.matchAll(/--([a-z][a-z0-9-]*)/g), (m) => m[1] ?? '');
        const declared = new Set(command.meta.options.map((opt) => opt.name));
        expect(flagsInSummary.filter((flag) => !declared.has(flag))).toEqual([]);
      });
    });
  }

  it('references no PHANTOM command names in any prose (summary / option descriptions / responseShape / example / bodyTemplate)', () => {
    // A write meta names read commands too (`find-mail-drafts`, `list-mail-attachments`):
    // every name of the directory is a real command of the split.
    const valid = new Set<string>([...Object.keys(commands), ...WRITE_LIFECYCLE_ENTRIES.map((entry) => entry.name), ...COMMAND_DIRECTORY.keys()]);
    const verbKebab = /^(get|list|convert|download|extract|search|resolve|find|create|update|delete|read|my|next|microsoft|scopes)-[a-z0-9-]+$/;
    // The phantom references in one command's prose.
    const phantomsIn = (name: string, prose: string): string[] => {
      const phantom: string[] = [];
      // (a) explicit `{<package>-bin} <name>` invocations: the directory must place <name> in that package.
      for (const m of prose.matchAll(/\{(?:(auth|read|write)-)?bin\}\s+([a-z][a-z0-9-]+)/g)) {
        const ref = m[2] ?? '';
        if (!valid.has(ref) || commandRole(ref) !== m[1]) phantom.push(`${name}: ${m[0]}`);
      }
      // (b) backticked command-shaped tokens (verb-prefixed kebab).
      for (const m of prose.matchAll(/`([a-z][a-z0-9-]+)`/g)) {
        const tok = m[1] ?? '';
        if (verbKebab.test(tok) && !valid.has(tok)) phantom.push(`${name}: \`${tok}\``);
      }
      return phantom;
    };
    const phantom = Object.entries(commands).flatMap(([name, command]) =>
      phantomsIn(
        name,
        [command.meta.summary, command.meta.example, command.meta.responseShape ?? '', command.meta.bodyTemplate ?? '', ...command.meta.options.map((o) => o.description)].join(
          '\n'
        )
      )
    );
    expect(phantom).toEqual([]);
  });
});
