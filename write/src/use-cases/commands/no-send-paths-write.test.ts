import { describe, expect, it } from 'bun:test';
import { sendsMail } from '@ask-marcel/office-core/domain/mail-send-guard.ts';
import { commands } from './write-commands.ts';

/*
 * No command of the write package names a Graph path that sends mail: the
 * write view would refuse it at run time (mail-send-guard.ts), so a command
 * whose `graphPathTemplate` names one is a defect caught here, before release.
 * Each template is read for every Graph path it names (a template can name an
 * alternative path in its prose, e.g. `, or /createReply when ...`).
 */

// The Graph paths a template names: each run of characters that starts with `/`.
const pathsOf = (template: string): ReadonlyArray<string> => template.match(/\/[^\s,()]+/g) ?? [];

describe('the Graph paths of the write commands', () => {
  it('name no path that sends mail', () => {
    const sending = Object.entries(commands).flatMap(([name, command]) =>
      pathsOf(command.meta.graphPathTemplate)
        .filter((path) => sendsMail(path))
        .map((path) => `${name}: ${path}`)
    );
    expect(sending).toEqual([]);
  });

  it('name at least one path each, so the check above reads every command', () => {
    const silent = Object.entries(commands)
      .filter(([, command]) => pathsOf(command.meta.graphPathTemplate).length === 0)
      .map(([name]) => name);
    expect(silent).toEqual([]);
  });

  it('would catch a template that names a sending path in its prose', () => {
    expect(pathsOf('/me/messages/{message-id}/createReply, or /send when {now} is true').filter((path) => sendsMail(path))).toEqual(['/send']);
  });
});
