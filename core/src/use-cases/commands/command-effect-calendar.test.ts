import { describe, expect, it } from 'bun:test';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { describeWrites, WRITE_EFFECTS } from './command-effect.ts';

/*
 * The write class of the calendar draft commands: an event in the calendar of
 * the user that has no attendees, so Graph sends no invitation. It deletes
 * nothing, so it is not destructive; the delete of a calendar draft is on the
 * delete class.
 */

describe('the write class of the calendar drafts', () => {
  it('describes what a command of the class does to the tenant', () => {
    expect(describeWrites({ 'create-widget-event': fakeCommand({ effect: 'calendar-draft' }) })).toBe(
      'create-widget-event: create or change a DRAFT event in your calendar: it has no attendees, so no invitation is sent (you add the attendees and send it yourself from Outlook)'
    );
  });

  it('is not destructive', () => {
    expect(WRITE_EFFECTS['calendar-draft'].destructive).toBe(false);
  });
});
