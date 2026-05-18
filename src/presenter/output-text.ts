type Cursors = {
  readonly nextLink?: string;
  readonly deltaLink?: string;
  readonly count?: number;
};

const HOIST_KEYS: ReadonlySet<string> = new Set(['@odata.nextLink', '@odata.deltaLink', '@odata.count']);

const isRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const isTextPayload = (record: Record<string, unknown>): boolean =>
  typeof record.contentType === 'string' && record.contentType.startsWith('text/') && typeof record.text === 'string';

const isBinaryPayload = (record: Record<string, unknown>): boolean =>
  typeof record.contentType === 'string' && typeof record.size === 'number' && typeof record.base64 === 'string';

const encodeScalar = (value: unknown): string => {
  if (value === null) return 'null';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
  return JSON.stringify(value);
};

const inlineArray = (arr: ReadonlyArray<unknown>): string => `[${arr.map(encodeScalar).join(', ')}]`;

const renderArrayField = (key: string, arr: ReadonlyArray<unknown>, indent: string): ReadonlyArray<string> => {
  if (arr.length === 0) return [`${indent}${key}: []`];
  if (arr.every((v) => !isRecord(v))) return [`${indent}${key}: ${inlineArray(arr)}`];
  const childIndent = `${indent}  `;
  const blocks: string[] = [`${indent}${key}:`];
  arr.forEach((item, idx) => {
    if (idx > 0) blocks.push('');
    blocks.push(...renderRecordLines(item as Record<string, unknown>, childIndent));
  });
  return blocks;
};

const renderRecordLines = (record: Record<string, unknown>, indent: string): ReadonlyArray<string> => {
  const lines: string[] = [];
  for (const [key, value] of Object.entries(record)) {
    if (value === undefined) continue;
    if (isRecord(value)) {
      lines.push(`${indent}${key}:`);
      lines.push(...renderRecordLines(value, `${indent}  `));
    } else if (Array.isArray(value)) {
      lines.push(...renderArrayField(key, value, indent));
    } else {
      lines.push(`${indent}${key}: ${encodeScalar(value)}`);
    }
  }
  return lines;
};

const extractCursors = (record: Record<string, unknown>): { readonly stripped: Record<string, unknown>; readonly cursors: Cursors } => {
  const stripped: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(record)) if (!HOIST_KEYS.has(key)) stripped[key] = value;
  const nextLink = record['@odata.nextLink'];
  const deltaLink = record['@odata.deltaLink'];
  const count = record['@odata.count'];
  return {
    stripped,
    cursors: {
      ...(typeof nextLink === 'string' ? { nextLink } : {}),
      ...(typeof deltaLink === 'string' ? { deltaLink } : {}),
      ...(typeof count === 'number' ? { count } : {}),
    },
  };
};

const renderFooter = (cursors: Cursors): string => {
  const parts: string[] = [];
  if (cursors.nextLink !== undefined) parts.push(`next: ${cursors.nextLink}`);
  if (cursors.deltaLink !== undefined) parts.push(`delta: ${cursors.deltaLink}`);
  if (cursors.count !== undefined) parts.push(`count: ${cursors.count}`);
  return parts.length === 0 ? '' : `--- ${parts.join(' · ')}`;
};

const appendFooter = (body: string, footer: string): string => (footer === '' ? `${body}\n` : `${body}\n\n${footer}\n`);

const renderCollection = (items: ReadonlyArray<unknown>, footer: string): string => {
  if (items.length === 0) return appendFooter('(no items)', footer);
  const blocks = items.map((item) => (isRecord(item) ? renderRecordLines(item, '').join('\n') : encodeScalar(item)));
  return appendFooter(blocks.join('\n\n'), footer);
};

const renderEnveloped = (data: Record<string, unknown>): string => {
  if (isTextPayload(data)) return `${data.text as string}\n`;
  if (isBinaryPayload(data)) return `binary: ${data.contentType as string}, ${data.size as number} bytes — use --output-path to save\n`;
  const { stripped, cursors } = extractCursors(data);
  const footer = renderFooter(cursors);
  if (Array.isArray(stripped.value) && Object.keys(stripped).length === 1) return renderCollection(stripped.value, footer);
  return appendFooter(renderRecordLines(stripped, '').join('\n'), footer);
};

const renderTextOutput = (data: unknown): string => {
  if (isRecord(data)) return renderEnveloped(data);
  if (Array.isArray(data)) return `${data.map((v) => (isRecord(v) ? renderRecordLines(v, '').join('\n') : encodeScalar(v))).join('\n')}\n`;
  return `${encodeScalar(data)}\n`;
};

export { renderTextOutput };
