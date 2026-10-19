// The write case of the meta suite: the scope the PDF converters declare. Moved out of
// meta.test.ts with its text unchanged (package split, step 16), so the write package takes
// this file whole.
import { describe, expect, it } from 'bun:test';
import { commands } from './write-commands.ts';

describe('command meta — invariants on every registered command', () => {
  it('declares Files.ReadWrite on each PDF converter, because it uploads a temporary file to the OneDrive to convert it', () => {
    for (const name of ['convert-mail-attachment-to-pdf', 'convert-calendar-event-attachment-to-pdf', 'convert-group-post-attachment-to-pdf']) {
      expect({ name, scopes: commands[name]?.meta.scopesRequired }).toEqual({ name, scopes: expect.arrayContaining(['Files.ReadWrite']) });
    }
  });
});
