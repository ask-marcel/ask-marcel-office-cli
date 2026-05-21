import { describe, expect, it } from 'bun:test';
import type { Command, CommandMeta } from './command-types.ts';
import { commands } from './index.ts';

const camelToKebab = (s: string): string => s.replaceAll(/([A-Z])/g, '-$1').toLowerCase();

const schemaKeys = (schema: Command['schema']): string[] => {
  const shape = (schema as unknown as { shape?: Record<string, unknown> }).shape;
  return shape ? Object.keys(shape) : [];
};

const placeholders = (template: string): string[] => Array.from(template.matchAll(/\{([a-z][a-z0-9-]*)\}/g), (m) => m[1] ?? '');

// True when a Zod field is `.optional()` — used by the `required` invariant
// below to keep meta.required in lock-step with the schema shape.
const isOptionalSchemaField = (schema: Command['schema'], key: string): boolean => {
  const shape = (schema as unknown as { shape?: Record<string, { isOptional?: () => boolean }> }).shape;
  const field = shape?.[key];
  return typeof field?.isOptional === 'function' ? field.isOptional() : false;
};

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

      it('references each per-command option at least once across graphPathTemplate + bodyTemplate, and references nothing else (runtime-additive query flags — OData + the chatsvcagg page-size/message-token analogues — are excluded)', () => {
        const runtimeFlagNames = new Set([
          'top',
          'skip',
          'select',
          'filter',
          'orderby',
          'expand',
          // chatsvcagg uses different query-param names than Graph's OData;
          // they're functionally identical (runtime-additive, not path
          // placeholders) so the invariant excludes them too.
          'page-size',
          // pre-2026-05 chatsvcagg name; kept here so the test isn't
          // a regression risk if someone reintroduces a similar flag
          // before the rename rolls everywhere.
          'skip-token',
          // post-2026-05 substrate pagination cursor for
          // list-teams-chat-messages.
          'message-token',
        ]);
        const expected = Array.from(new Set(cmd.meta.options.filter((o) => !runtimeFlagNames.has(o.name)).map((o) => o.name))).toSorted((a, b) => a.localeCompare(b));
        // `{region}` is an infra-level placeholder on the post-2026-05
        // chatsvcagg substrate (`teams.microsoft.com/api/csa/{region}/...`).
        // It's resolved from cached auth state, NEVER user-supplied, so it's
        // not in `options` — strip it before checking so the invariant
        // still enforces "every other placeholder maps to an option".
        const combined = `${cmd.meta.graphPathTemplate} ${cmd.meta.bodyTemplate ?? ''}`.replaceAll('{region}', '');
        const found = Array.from(new Set(placeholders(combined))).toSorted((a, b) => a.localeCompare(b));
        expect(found).toEqual(expected);
      });

      it('publishes a Microsoft Learn URL for the underlying Graph endpoint or guide', () => {
        expect(cmd.meta.graphDocsUrl).toMatch(/^https:\/\/learn\.microsoft\.com\/en-us\/graph\//);
      });

      it('provides a runnable example beginning with `ask-marcel`', () => {
        expect(cmd.meta.example).toMatch(/^ask-marcel /);
      });

      it('has a non-empty category, graphMethod from the HTTP-verb set, non-empty graphPathTemplate, and non-empty responseShape when declared (invariants on the meta block itself, killing inert-config-string mutants)', () => {
        expect(cmd.meta.category.trim().length).toBeGreaterThan(0);
        expect(['GET', 'POST', 'PUT', 'DELETE', 'PATCH']).toContain(cmd.meta.graphMethod);
        expect(cmd.meta.graphPathTemplate.trim().length).toBeGreaterThan(0);
        // responseShape is optional on the type but, when present, must
        // not be the empty string — that's the mutant we want to kill.
        if (cmd.meta.responseShape !== undefined) {
          expect(cmd.meta.responseShape.trim().length).toBeGreaterThan(0);
        }
      });

      it('pairs each option.required with the schema field optionality — required:true iff the Zod field is NOT .optional() — keeping the help text and the runtime validator in lock-step', () => {
        for (const opt of cmd.meta.options) {
          const isOpt = isOptionalSchemaField(cmd.schema, opt.key);
          expect({ key: opt.key, required: opt.required }).toEqual({ key: opt.key, required: !isOpt });
        }
      });

      it('summary references only flags that are actually registered as options or aliases on this command', () => {
        const flagsInSummary = Array.from(cmd.meta.summary.matchAll(/--([a-z][a-z0-9-]*)/g), (m) => m[1] ?? '');
        if (flagsInSummary.length === 0) return;
        const declared = new Set<string>();
        for (const opt of cmd.meta.options) {
          declared.add(opt.name);
          for (const alias of opt.aliases ?? []) declared.add(alias.name);
        }
        for (const flag of flagsInSummary) expect(declared).toContain(flag);
      });
    });
  }
});
