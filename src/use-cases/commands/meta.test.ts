import { describe, expect, it } from 'bun:test';
import type { Command, CommandMeta } from './command-types.ts';
import { commands } from './index.ts';

const camelToKebab = (s: string): string => s.replaceAll(/([A-Z])/g, '-$1').toLowerCase();

const schemaKeys = (schema: Command['schema']): string[] => {
  const shape = (schema as unknown as { shape?: Record<string, unknown> }).shape;
  return shape ? Object.keys(shape) : [];
};

const placeholders = (template: string): string[] => Array.from(template.matchAll(/\{([a-z][a-z0-9-]*)\}/g), (m) => m[1] ?? '');

type PopulatedEntry = readonly [string, Command & { meta: CommandMeta }];

const populated: ReadonlyArray<PopulatedEntry> = Object.entries(commands).map(([name, cmd]) => [name, cmd as Command & { meta: CommandMeta }] as PopulatedEntry);

describe('command meta — invariants on every registered command', () => {
  it('every command in the registry has a meta block', () => {
    for (const [name, cmd] of Object.entries(commands)) {
      expect({ name, hasMeta: cmd.meta !== undefined }).toEqual({ name, hasMeta: true });
    }
  });

  for (const [name, cmd] of populated) {
    describe(`meta for \`${name}\``, () => {
      it('has a non-empty summary', () => {
        expect(cmd.meta.summary.trim().length).toBeGreaterThan(0);
      });

      it('declares one CommandOptionMeta entry per Zod schema field', () => {
        const keys = schemaKeys(cmd.schema).toSorted((a, b) => a.localeCompare(b));
        const optionKeys = cmd.meta.options.map((o) => o.key).toSorted((a, b) => a.localeCompare(b));
        expect(optionKeys).toEqual(keys);
      });

      it('uses kebab-case `name` matching the camelCase `key` in every option', () => {
        for (const opt of cmd.meta.options) {
          expect(opt.name).toBe(camelToKebab(opt.key));
          expect(opt.description.trim().length).toBeGreaterThan(0);
        }
      });

      it('references each option at least once across graphPathTemplate + bodyTemplate, and references nothing else', () => {
        const expected = Array.from(new Set(cmd.meta.options.map((o) => o.name))).toSorted((a, b) => a.localeCompare(b));
        const combined = `${cmd.meta.graphPathTemplate} ${cmd.meta.bodyTemplate ?? ''}`;
        const found = Array.from(new Set(placeholders(combined))).toSorted((a, b) => a.localeCompare(b));
        expect(found).toEqual(expected);
      });

      it('publishes a Microsoft Learn URL for the underlying Graph endpoint or guide', () => {
        expect(cmd.meta.graphDocsUrl).toMatch(/^https:\/\/learn\.microsoft\.com\/en-us\/graph\//);
      });

      it('provides a runnable example beginning with `ask-marcel`', () => {
        expect(cmd.meta.example).toMatch(/^ask-marcel /);
      });
    });
  }
});
