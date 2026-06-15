# Tasks commands

## get-planner-bucket
Get the metadata of a single Microsoft Planner bucket (column / lane).
Required: --planner-bucket-id
Example: ask-marcel get-planner-bucket --planner-bucket-id 'sFNeQRFu_kqhxpwwAhmA15gAGfoT'
Graph: GET /planner/buckets/{planner-bucket-id}

## get-planner-plan
Get the metadata of a single Microsoft Planner plan (title, owner group, container).
Required: --planner-plan-id
Example: ask-marcel get-planner-plan --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'
Graph: GET /planner/plans/{planner-plan-id}

## get-planner-task
Get the metadata of a single Microsoft Planner task (title, assignees, dates, completion).
Required: --planner-task-id
Example: ask-marcel get-planner-task --planner-task-id '01tx7Ic7-USXEwt0lvR1cmgAH8gK'
Graph: GET /planner/tasks/{planner-task-id}

## get-planner-task-details
Get the rich details (description, checklist, references) of a Microsoft Planner task.
Required: --planner-task-id
Example: ask-marcel get-planner-task-details --planner-task-id '01tx7Ic7-USXEwt0lvR1cmgAH8gK'
Graph: GET /planner/tasks/{planner-task-id}/details

## get-todo-task
Get a single Microsoft To Do task by its ID and its parent list ID. Use `--select` to slim the response (e.g. `--select id,status`) or `--expand checklistItems` / `--expand linkedResources` to inline child collections. Known Graph quirk: any `--select` combo that includes `title` trips `RequestBroke
Required: --todo-task-list-id --todo-task-id
Optional: --select --expand
Example: ask-marcel get-todo-task --todo-task-list-id 'AAMkAGI...' --todo-task-id 'AAMkABC...'
Graph: GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}

## list-incomplete-planner-tasks
List every incomplete Microsoft Planner task assigned to or owned by the signed-in user, across every plan. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-percent predicate, an
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-incomplete-planner-tasks --top 25
Graph: GET /me/planner/tasks?$filter=percentComplete ne 100

## list-incomplete-todo-tasks
List every incomplete Microsoft To Do task in a given list (status not equal to `completed`). Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-status predicate, and Graph rejects
Required: --todo-task-list-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-incomplete-todo-tasks --todo-task-list-id 'tasks' --top 5
Graph: GET /me/todo/lists/{todo-task-list-id}/tasks?$filter=status ne 'completed'

## list-plan-buckets
List the buckets (columns / lanes) of a Microsoft Planner plan. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — slice / sort client-side.
Required: --planner-plan-id
Optional: --select
Example: ask-marcel list-plan-buckets --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'
Graph: GET /planner/plans/{planner-plan-id}/buckets

## list-plan-tasks
List every task within a Microsoft Planner plan, regardless of completion status (Graph orders by `orderHint`). Use `list-incomplete-planner-tasks` for the across-plans incomplete view. Note: Graph silently ignores standard OData query parameters on `/planner/plans/{id}/tasks` (`$top` returns the fu
Required: --planner-plan-id
Example: ask-marcel list-plan-tasks --planner-plan-id 'xqQg5FS2LkCp935s-FIFm5gAB6'
Graph: GET /planner/plans/{planner-plan-id}/tasks

## list-planner-plans
List every Microsoft Planner plan the signed-in user has access to (across every group). Use this to discover plan IDs without needing an existing task as the entry point. Note: Graph silently drops `$top`, `$skip`, `$filter`, and `$orderby` on this endpoint, so the CLI advertises only `--select` — 
Optional: --select
Example: ask-marcel list-planner-plans
Graph: GET /me/planner/plans

## list-planner-tasks
List every Microsoft Planner task assigned to or owned by the signed-in user, across all plans.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-planner-tasks
Graph: GET /me/planner/tasks

## list-todo-linked-resources
List the linked resources (URLs, emails, files) attached to a Microsoft To Do task.
Required: --todo-task-list-id --todo-task-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-todo-linked-resources --todo-task-list-id 'AAMkAGI...' --todo-task-id 'AAMkABC...'
Graph: GET /me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}/linkedResources

## list-todo-task-lists
List the signed-in user's Microsoft To Do task lists (e.g. `Tasks`, `Flagged Emails`, custom lists). Note: Graph rejects `$select` and `$orderby` on this endpoint with `RequestBroker--ParseUri`, so the CLI does not expose those flags — slice / sort client-side.
Optional: --top --skip --filter --expand
Example: ask-marcel list-todo-task-lists
Graph: GET /me/todo/lists

## list-todo-tasks
List every task in a single Microsoft To Do task list, regardless of completion status. Use `list-incomplete-todo-tasks` if you only want the open ones. Known Graph quirk: certain `--select` combinations (notably any combo that includes `title`) trip `RequestBroker--ParseUri` on this endpoint; the C
Required: --todo-task-list-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-todo-tasks --todo-task-list-id 'AAMkAGI...'
Graph: GET /me/todo/lists/{todo-task-list-id}/tasks

## list-todo-tasks-delta
Track incremental task changes (added / updated / completed / deleted) within a single Microsoft To Do list. The first call returns the current snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Note: Graph rejects standard OData query parameters on
Required: --todo-task-list-id
Example: ask-marcel list-todo-tasks-delta --todo-task-list-id 'AAMkAD...'
Graph: GET /me/todo/lists/{todo-task-list-id}/tasks/delta()
