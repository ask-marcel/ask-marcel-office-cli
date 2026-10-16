import { describe, expect, it } from 'bun:test';
import { closestNames, didYouMean } from './closest-names.ts';

const COMMANDS = ['list-teams-chat-messages', 'list-teams-chat-history', 'list-teams-chats-with-messages', 'list-mail-messages', 'get-mail-message', 'list-chat-members'];

describe('the names closest to a mistyped one', () => {
  it('finds a command whose words contain the guessed words in order, fewest extra words first', () => {
    expect(closestNames('list-chat-messages', COMMANDS)).toEqual(['list-teams-chat-messages', 'list-teams-chats-with-messages']);
  });

  it('finds a flag the same way, ignoring the leading dashes, a doubled dash and letter case', () => {
    expect(closestNames('--List-Id', ['--top', '--skip', '--todo-task-list-id', '--select'])).toEqual(['--todo-task-list-id']);
    expect(closestNames('list--id', ['--top', '--todo-task-list-id'])).toEqual(['--todo-task-list-id']);
  });

  it('catches a typo within three edits when no word order matches', () => {
    expect(closestNames('list-mail-mesages', COMMANDS)).toEqual(['list-mail-messages']);
    expect(closestNames('get-mail-msge', COMMANDS)).toEqual(['get-mail-message']);
    expect(closestNames('get-mail-mge', COMMANDS)).toEqual([]);
  });

  it('ranks word matches before typos, breaks ties by the shorter name, then by the order given, and keeps three at most', () => {
    const names = ['list-mail-folders', 'list-mail-folder-messages', 'list-mail', 'list-mail-items', 'list-mali'];
    expect(closestNames('list-mail', names)).toEqual(['list-mail', 'list-mail-items', 'list-mail-folders']);
    expect(closestNames('list-mail', ['list-mals', 'list-mail-folder-messages'])).toEqual(['list-mail-folder-messages', 'list-mals']);
    expect(closestNames('list-mial', ['list-mail-folders', 'list-mali', 'list-mail'])).toEqual(['list-mali', 'list-mail']);
  });

  it('answers nothing for a name unlike every candidate, or for an empty one', () => {
    expect(closestNames('frobnicate', COMMANDS)).toEqual([]);
    expect(closestNames('--', COMMANDS)).toEqual([]);
  });
});

describe('the did-you-mean sentence', () => {
  it('lists one, two or three names in backticks, led by a space so it can follow a sentence', () => {
    expect(didYouMean('list-mail-mesages', COMMANDS)).toBe(' Did you mean `list-mail-messages`?');
    expect(didYouMean('list-chat-messages', COMMANDS)).toBe(' Did you mean `list-teams-chat-messages` or `list-teams-chats-with-messages`?');
    expect(didYouMean('list-mail', ['list-mail-folders', 'list-mail', 'list-mail-items'])).toBe(' Did you mean `list-mail`, `list-mail-items` or `list-mail-folders`?');
  });

  it('takes its own lead-in, and says nothing when no name is close', () => {
    expect(didYouMean('--list-id', ['--todo-task-list-id'], 'For `--list-id`, did you mean')).toBe(' For `--list-id`, did you mean `--todo-task-list-id`?');
    expect(didYouMean('frobnicate', COMMANDS)).toBe('');
  });
});
