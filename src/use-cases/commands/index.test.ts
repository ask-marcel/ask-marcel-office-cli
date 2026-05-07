import { describe, expect, it } from 'bun:test';
import { commands } from './index.ts';

describe('commands index', () => {
  it('exports a non-empty commands record', () => {
    const names = Object.keys(commands);
    expect(names.length).toBeGreaterThan(0);
    expect(names).toContain('list-drives');
    expect(names).toContain('list-joined-teams');
  });

  it('does not register `list-sharepoint-site-items` — Graph has no list-less site/items collection endpoint', () => {
    expect(Object.keys(commands)).not.toContain('list-sharepoint-site-items');
  });

  it('points `get-sharepoint-site-item --item-id` at the two-step discovery chain via `list-sharepoint-site-lists` and `list-sharepoint-site-list-items`', () => {
    const cmd = commands['get-sharepoint-site-item'];
    if (cmd === undefined) throw new Error('get-sharepoint-site-item is missing from the registry');
    const itemIdOption = cmd.meta.options.find((o) => o.key === 'itemId');
    if (itemIdOption === undefined) throw new Error('itemId option is missing from get-sharepoint-site-item');
    expect(itemIdOption.description).toContain('list-sharepoint-site-lists');
    expect(itemIdOption.description).toContain('list-sharepoint-site-list-items');
  });
});
