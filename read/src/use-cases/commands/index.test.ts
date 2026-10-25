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

  it("does not register `get-sharepoint-site-item` — Graph's /sites/{site-id}/items/{item-id} expects a SharePoint UNIQUEID that no other manifest command produces, so the chain dead-ended", () => {
    expect(Object.keys(commands)).not.toContain('get-sharepoint-site-item');
  });
});
