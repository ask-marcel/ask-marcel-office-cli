import { describe, expect, it } from 'bun:test';
import { meta as updatePlannerTaskDetails } from './update-planner-task-details.ts';
import { meta as addPlannerChecklistItem } from './add-planner-checklist-item.ts';
import { meta as updatePlannerChecklistItem } from './update-planner-checklist-item.ts';
import { meta as removePlannerChecklistItem } from './remove-planner-checklist-item.ts';

/*
 * The docs of the details and checklist commands of Planner, as the help, the docs
 * pages and the MCP tools show them.
 */

describe('the docs of the Planner details and checklist commands', () => {
  it('pins the meta of update-planner-task-details', () => {
    expect(updatePlannerTaskDetails).toStrictEqual({
      summary:
        'Change the details of a Planner task: the description as plain text (--description) or the notes as HTML (--notes-html, on Graph beta), the preview of the card, and the checklist from a file (--checklist-file). The undo writes the details of before back.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/tasks/{task-id}/details',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'description',
          key: 'description',
          required: false,
          description: 'The description as plain text. It replaces the description and the notes of before. Not with --notes-html.',
        },
        {
          name: 'notes-html',
          key: 'notesHtml',
          required: false,
          description:
            'The notes as HTML, sent to Graph beta (v1.0 has no notes). Planner keeps their plain text as the description. Not with --description: Graph refuses both in one request.',
        },
        {
          name: 'checklist-file',
          key: 'checklistFile',
          required: false,
          description:
            'A local UTF-8 text file with the checklist: one item on each line, at most 20 items of at most 100 characters, "[x] " before a checked item and "[ ] " before an item that is not checked. An item of the task with the same title stays, the other items of the task are removed, and the new titles are added. At most 64000 bytes. A relative path starts from the folder the CLI runs in.',
        },
        {
          name: 'preview',
          key: 'preview',
          required: false,
          description: 'What the card of the task shows on the board.',
          argumentHint: {
            kind: 'magicValue',
            values: ['automatic', 'noPreview', 'checklist', 'description', 'reference'],
          },
        },
      ],
      example: '{write-bin} update-planner-task-details --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --notes-html "<p>See the <b>budget</b></p>" --checklist-file steps.txt --dry-run',
      bodyTemplate:
        "{ description?: '{description}' | notes?: { contentType: 'html', content: '{notes-html}' }, previewType?: '{preview}', checklist?: { <key>: { title, isChecked } | null } (the items of {checklist-file}) } with If-Match: <ETag of the details>",
      effect: 'task-write',
      undo: 'writes the description, the preview and the checklist of before back (a removed item comes back with a new id), only if the values that the run set did not change. HTML notes come back as their plain text',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of add-planner-checklist-item', () => {
    expect(addPlannerChecklistItem).toStrictEqual({
      summary: 'Add an item to the checklist of a Planner task (at most 20 items). The undo removes the item.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/tasks/{task-id}/details',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'title',
          key: 'title',
          required: true,
          description: 'The title of the item, at most 100 characters. The checklist must not have an item of this title.',
        },
        {
          name: 'checked',
          key: 'checked',
          required: false,
          description: '`true` checks the item, `false` unchecks it.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} add-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --title "Book the room" --dry-run',
      bodyTemplate: "{ checklist: { <new key>: { title: '{title}', isChecked: {checked} } } } with If-Match: <ETag of the details>",
      effect: 'task-write',
      undo: 'removes the item, only if its title and its state did not change',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of update-planner-checklist-item', () => {
    expect(updatePlannerChecklistItem).toStrictEqual({
      summary: 'Change the title of an item of the checklist of a Planner task, or check or uncheck it. The undo writes the title and the state of before back.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/tasks/{task-id}/details',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'item',
          key: 'item',
          required: true,
          description: 'The item of the checklist: its id, as get-planner-task-details returns it, or its exact title. It must name exactly one item.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'title',
          key: 'title',
          required: false,
          description: 'The new title of the item, at most 100 characters.',
        },
        {
          name: 'checked',
          key: 'checked',
          required: false,
          description: '`true` checks the item, `false` unchecks it.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} update-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --item "Book the room" --checked true --dry-run',
      bodyTemplate: "{ checklist: { <key of {item}>: { title: '{title}', isChecked: {checked} } } } with If-Match: <ETag of the details>",
      effect: 'task-write',
      undo: 'writes the title and the state of before back, only if the title and the state that the run set did not change',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of remove-planner-checklist-item', () => {
    expect(removePlannerChecklistItem).toStrictEqual({
      summary: 'Remove an item from the checklist of a Planner task. The undo adds the item again with a new id.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/tasks/{task-id}/details',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'item',
          key: 'item',
          required: true,
          description: 'The item of the checklist: its id, as get-planner-task-details returns it, or its exact title. It must name exactly one item.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
      ],
      example: '{write-bin} remove-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --item "Book the room" --dry-run',
      bodyTemplate: '{ checklist: { <key of {item}>: null } } with If-Match: <ETag of the details>',
      effect: 'task-write',
      undo: 'adds the item again with its title, its state and a new id, only if it is still not in the checklist',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });
});
