/**
 * Shared markdown render primitives for the per-format metadata blocks
 * (docx, xlsx, …). Sections with no entries emit `_(none)_` so the output
 * stays grep-stable — an LLM can ask "is there a Comments section?" and get
 * a yes/no answer without ambiguity.
 *
 * Pipe characters in cell values are escaped as `\|` so the rendered tables
 * stay valid even when property values or comment bodies contain literal pipes.
 */

const NONE = '_(none)_';

const escapeCell = (s: string): string => s.replace(/\|/g, '\\|').replace(/\r?\n/g, ' ').trim();

const renderKv = (props: Readonly<Record<string, string>>): string => {
  const entries = Object.entries(props).filter(([, v]) => v !== '');
  if (entries.length === 0) return NONE;
  return entries.map(([k, v]) => `- **${k}**: ${escapeCell(v)}`).join('\n');
};

const renderTable = (rows: ReadonlyArray<ReadonlyArray<string>>, headers: ReadonlyArray<string>): string => {
  if (rows.length === 0) return NONE;
  const head = `| ${headers.join(' | ')} |`;
  const sep = `| ${headers.map(() => '---').join(' | ')} |`;
  const body = rows.map((r) => `| ${r.map(escapeCell).join(' | ')} |`).join('\n');
  return [head, sep, body].join('\n');
};

const renderBullets = (items: ReadonlyArray<string>): string => {
  if (items.length === 0) return NONE;
  return items.map((i) => `- ${escapeCell(i)}`).join('\n');
};

export { escapeCell, NONE, renderBullets, renderKv, renderTable };
