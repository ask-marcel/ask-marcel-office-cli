import { describe, expect, it } from 'bun:test';
import { meta as addPlannerTaskFile } from './add-planner-task-file.ts';
import { meta as addPlannerTaskLink } from './add-planner-task-link.ts';
import { meta as removePlannerTaskLink } from './remove-planner-task-link.ts';
import { meta as updatePlannerPlan } from './update-planner-plan.ts';

/*
 * The docs of the link, file and plan commands of Planner, as the help, the docs
 * pages and the MCP tools show them.
 */

describe('the docs of the Planner reference and plan commands', () => {
  it('pins the meta of add-planner-task-file', () => {
    expect(addPlannerTaskFile).toStrictEqual({
      summary:
        'Attach a local file to a Planner task of a plan that is not of a group: the file goes into the folder "Planner attachments" of your OneDrive (shared with nobody), and the task gets a reference to it. A plan of a group is refused. The undo removes the reference and moves the file to the recycle bin.',
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
          name: 'file',
          key: 'file',
          required: true,
          description:
            'A local file of at most 25000000 bytes. It goes into the folder "Planner attachments" of your OneDrive under its own name; a name that the folder has already is refused. A relative path starts from the folder the CLI runs in.',
        },
        {
          name: 'alias',
          key: 'alias',
          required: false,
          description: 'The name of the reference that Planner shows, at most 255 characters. The default is the name of the file.',
        },
      ],
      example: '{write-bin} add-planner-task-file --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --file ./budget.xlsx --dry-run',
      bodyTemplate:
        "PUT /me/drive/root:/Planner attachments/<name of {file}>:/content, then { references: { <webUrl of the file, encoded>: { alias: '{alias}', type: 'Word' | 'Excel' | 'PowerPoint' | 'Other' } } } with If-Match: <ETag of the details>",
      effect: 'task-write',
      undo: 'removes the reference, only if its name and its type did not change, and moves the file to the recycle bin of your OneDrive, only if its content did not change',
      scopesRequired: ['Tasks.ReadWrite', 'Files.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of add-planner-task-link', () => {
    expect(addPlannerTaskLink).toStrictEqual({
      summary: 'Add a link (an http or https URL, with a name) to the references of a Planner task. The undo removes the link.',
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
          name: 'url',
          key: 'url',
          required: true,
          description: 'The full http or https URL of the link, at most 2048 characters.',
        },
        {
          name: 'alias',
          key: 'alias',
          required: false,
          description: 'The name of the link that Planner shows, at most 255 characters.',
        },
      ],
      example: '{write-bin} add-planner-task-link --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --url "https://example.com/spec" --alias "Spec" --dry-run',
      bodyTemplate: "{ references: { <{url} with %, ., :, @ and # encoded>: { alias: '{alias}', type: 'Other' } } } with If-Match: <ETag of the details>",
      effect: 'task-write',
      undo: 'removes the link, only if its name and its type did not change',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of remove-planner-task-link', () => {
    expect(removePlannerTaskLink).toStrictEqual({
      summary: 'Remove a link or a file from the references of a Planner task (a file stays where it is). The undo adds the reference again.',
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
          name: 'url',
          key: 'url',
          required: true,
          description: 'The full http or https URL of the link, at most 2048 characters.',
        },
      ],
      example: '{write-bin} remove-planner-task-link --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --url "https://example.com/spec" --dry-run',
      bodyTemplate: '{ references: { <{url} with %, ., :, @ and # encoded>: null } } with If-Match: <ETag of the details>',
      effect: 'task-write',
      undo: 'adds the reference again with its name and its type, only if it is still not in the references',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of update-planner-plan', () => {
    expect(updatePlannerPlan).toStrictEqual({
      summary: 'Rename a Planner plan (--name) and its labels (--label-names). No command creates a plan. The undo writes the name and the label names of before back.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/plans/<id of {plan}> and /planner/plans/<id of {plan}>/details',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerplan-update',
      options: [
        {
          name: 'plan',
          key: 'plan',
          required: true,
          description: 'The plan: its name (no case) or its id, as list-planner-plans returns them.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'name',
          key: 'name',
          required: false,
          description: 'The new name of the plan, at most 255 characters.',
        },
        {
          name: 'label-names',
          key: 'labelNames',
          required: false,
          description:
            'The names of labels, separated by commas, in the form category1=Red,category5=Blocked (N from 1 to 25). An empty name removes the name: category2=. The other labels keep their names.',
        },
      ],
      example: '{write-bin} update-planner-plan --plan "Launch" --name "Launch 2027" --label-names "category1=Urgent,category2=" --dry-run',
      bodyTemplate:
        "{ title: '{name}' } on the plan, and { categoryDescriptions: { categoryN: '<name>' | null } } (from {label-names}) on its details, each with If-Match: <its ETag>",
      effect: 'task-write',
      undo: 'writes the name and the label names of before back, only if the values that the run set did not change',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });
});
