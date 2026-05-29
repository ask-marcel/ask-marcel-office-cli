import { describe, expect, it } from 'bun:test';
import { escapeCell, NONE, renderBullets, renderKv, renderMacros, renderTable } from './ooxml-metadata-to-markdown.ts';

describe('escapeCell', () => {
  it('escapes pipes so a value cannot break out of a markdown table cell', () => {
    expect(escapeCell('a|b')).toBe('a\\|b');
  });

  it('flattens CR/LF to single spaces and trims the result', () => {
    expect(escapeCell('  line one\r\nline two\n  ')).toBe('line one line two');
  });

  it('returns a plain value unchanged (aside from trimming)', () => {
    expect(escapeCell(' hello ')).toBe('hello');
  });
});

describe('NONE placeholder', () => {
  it('is the grep-stable `_(none)_` marker', () => {
    expect(NONE).toBe('_(none)_');
  });
});

describe('renderKv', () => {
  it('returns NONE for an empty record', () => {
    expect(renderKv({})).toBe('_(none)_');
  });

  it('renders one `- **key**: value` line per non-empty entry, dropping empty values', () => {
    expect(renderKv({ creator: 'Vincent', title: 'Q4', empty: '' })).toBe('- **creator**: Vincent\n- **title**: Q4');
  });

  it('escapes pipes in values', () => {
    expect(renderKv({ k: 'a|b' })).toBe('- **k**: a\\|b');
  });
});

describe('renderTable', () => {
  it('returns NONE for zero rows', () => {
    expect(renderTable([], ['a', 'b'])).toBe('_(none)_');
  });

  it('renders a GFM header + separator + one line per row', () => {
    expect(
      renderTable(
        [
          ['1', '2'],
          ['3', '4'],
        ],
        ['a', 'b']
      )
    ).toBe('| a | b |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |');
  });

  it('escapes pipes inside cells', () => {
    expect(renderTable([['x|y']], ['h'])).toBe('| h |\n| --- |\n| x\\|y |');
  });
});

describe('renderBullets', () => {
  it('returns NONE for an empty list', () => {
    expect(renderBullets([])).toBe('_(none)_');
  });

  it('renders one `- item` line per entry, escaping pipes', () => {
    expect(renderBullets(['first', 'a|b'])).toBe('- first\n- a\\|b');
  });
});

describe('renderMacros', () => {
  it('returns NONE when there are no macro parts', () => {
    expect(renderMacros([])).toBe('_(none)_');
  });

  it('flags each vbaProject.bin part with a code-execution warning', () => {
    expect(renderMacros(['word/vbaProject.bin'])).toBe('- `word/vbaProject.bin` — embedded VBA macro project; this file can execute code when opened, treat as untrusted');
  });
});
