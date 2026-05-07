import type { z } from 'zod';

const camelToKebab = (s: string): string => s.replaceAll(/([A-Z])/g, '-$1').toLowerCase();

const humanize = (issue: z.core.$ZodIssue): string => {
  if (issue.code === 'invalid_type') return 'is missing';
  if (issue.code === 'too_small') {
    const minimum = (issue as { minimum?: number | bigint }).minimum;
    if (typeof minimum === 'number' && minimum === 1) return 'is empty';
  }
  return issue.message;
};

const renderPath = (path: ReadonlyArray<PropertyKey>): string => path.map((p) => (typeof p === 'string' || typeof p === 'number' ? String(p) : '')).join('.');

/**
 * Render a Zod validation error as a single human-readable line keyed off the
 * CLI flag name. The audit's gripe with the previous formatter
 * (`messageRuleId: Too small: expected string to have >=1 characters`) was
 * that an LLM piping the JSON envelope through jq sees Zod jargon plus the
 * camelCase schema key — neither matches what the user typed on the command
 * line. This formatter prepends `--<kebab-case-flag>` and translates the two
 * common Zod codes to plain English: missing values become `is missing`,
 * empty strings become `is empty`. Multi-issue failures join with `; ` to
 * stay on one line.
 */
const formatZodError = (error: z.ZodError): string => {
  if (error.issues.length === 0) return 'validation failed';
  return error.issues
    .map((issue) => {
      if (issue.path.length === 0) return `<root>: ${issue.message}`;
      const path = renderPath(issue.path);
      const flag = `--${camelToKebab(path)}`;
      return `${flag} ${humanize(issue)}`;
    })
    .join('; ');
};

export { formatZodError };
