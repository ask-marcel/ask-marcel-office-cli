import { describe, expect, it } from 'bun:test';
import { renderTextOutput } from './output-text.ts';

describe('text output of a list that holds lists', () => {
  it('writes each inner list as JSON on the line of its field', () => {
    expect(
      renderTextOutput({
        id: 'r1',
        cells: [
          ['A1', 'B1'],
          [1, 2],
        ],
      })
    ).toBe('id: r1\ncells: [["A1","B1"], [1,2]]\n');
  });
});
