# Calendar commands

## convert-calendar-event-attachment-to-markdown
Convert an attachment on an Outlook calendar event to markdown. Polymorphic on the attachment’s `@odata.type` (shares the mail-attachment pipeline): fileAttachment decodes the inline bytes and runs them through the local conversion pipeline (docx via mammoth, xlsx via sheetjs, csv as markdown table,
Required: --event-id --attachment-id
Optional: --include-metadata
Example: ask-marcel convert-calendar-event-attachment-to-markdown --event-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'
Graph: GET /me/events/{event-id}/attachments/{attachment-id}

## convert-calendar-event-attachment-to-pdf
Convert an attachment on an Outlook calendar event to PDF on the fly (shares the mail-attachment pipeline). fileAttachment uploads the bytes to a temp folder under /me/drive, runs Graph `?format=pdf`, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and converts
Required: --event-id --attachment-id
Example: ask-marcel convert-calendar-event-attachment-to-pdf --event-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-path ./deck.pdf
Graph: GET /me/events/{event-id}/attachments/{attachment-id}

## get-calendar-event
Fetch a single calendar event by ID from the signed-in user’s default calendar. Pass `--select` to project only the fields you need (the full event body can be large with HTML body and attendee lists).
Required: --event-id
Optional: --select --expand
Example: ask-marcel get-calendar-event --event-id 'AAMkAGI2THVS...' --select id,subject,start,end,attendees
Graph: GET /me/events/{event-id}

## get-my-calendar
Return metadata for the signed-in user's *primary* calendar — `id`, `name`, `color`, `owner`, `canShare`, `canViewPrivateItems`, `canEdit`, `defaultOnlineMeetingProvider`. Sibling to `list-calendars` which returns every calendar (incl. shared / subscribed). Use `--select` to fetch only the fields yo
Optional: --select --expand
Example: ask-marcel get-my-calendar --select 'id,name'
Graph: GET /me/calendar

## get-specific-calendar-event
Fetch a single calendar event by ID from a specific calendar. `--calendar-id primary` (or `default`) targets the signed-in user's default calendar. Use `--select` to slim large event payloads (a typical event with body+attendees runs >50 KB).
Required: --calendar-id --event-id
Optional: --select --expand
Example: ask-marcel get-specific-calendar-event --calendar-id 'primary' --event-id 'AAMkABC...' --select 'id,subject,start,end'
Graph: GET /me/calendars/{calendar-id}/events/{event-id}

## list-calendar-event-attachments
List the attachments (file, item, reference) on a single Outlook calendar event. Ships an opinionated default `--select=id,name,contentType,size,isInline` so an LLM doesn't accidentally pull multi-MB `contentBytes` for every attachment. The `@odata.type` discriminator is always returned by Graph reg
Required: --event-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-event-attachments --event-id 'AAMkAGI2...'
Graph: GET /me/events/{event-id}/attachments

## list-calendar-event-instances
List the individual occurrences of a recurring calendar event over a date range. Both ISO date-time params are required by Graph. `--calendar-id` is optional and defaults to `primary` (the signed-in user’s default calendar) — most callers know the event-id but not which calendar it lives in. Pass an
Required: --event-id --start-date-time --end-date-time
Optional: --calendar-id --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-event-instances --calendar-id 'AAMkAGI2THVS...' --event-id 'AAMkABC...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'
Graph: GET /me/calendars/{calendar-id}/events/{event-id}/instances?startDateTime={start-date-time}&endDateTime={end-date-time}

## list-calendar-events
List the events in the signed-in user’s default calendar (does not expand recurrences).
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-events
Graph: GET /me/events

## list-calendar-events-delta
Get the incremental change set (added / modified / deleted events) for the signed-in user's default calendar. Use the `@odata.deltaLink` from a previous response to resume. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally; `$top` as a URL query is rejected by Graph
Optional: --top
Example: ask-marcel list-calendar-events-delta --top 50
Graph: GET /me/events/delta()

## list-calendar-group-calendars
List the calendars inside one calendar group.
Required: --calendar-group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-group-calendars --calendar-group-id 'AAMkADk0...'
Graph: GET /me/calendarGroups/{calendar-group-id}/calendars

## list-calendar-groups
List the signed-in user's calendar groups — Outlook's organizational layer above individual calendars (e.g. "My Calendars", "Other Calendars", "Birthdays"). Use the returned `id` with `list-calendar-group-calendars` to drill in.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-groups
Graph: GET /me/calendarGroups

## list-calendar-view
List the signed-in user's default-calendar events with recurrence expanded into individual occurrences in a date range. Both date-time params accept strict ISO 8601 (`2026-04-01T00:00:00Z`) AND the CLI's relative shapes (`7d`, `today`, `monday`, `start-of-month`, …) so a question like "what's on my 
Required: --start-date-time --end-date-time
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendar-view --start-date-time 'start-of-week' --end-date-time 'end-of-week'
Graph: GET /me/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}

## list-calendar-view-delta
Get the first page of the incremental change set of expanded calendar-view occurrences over a date range. Subsequent pages: feed the returned `@odata.nextLink` to `next-page`; resume later via the `@odata.deltaLink`. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally
Required: --start-date-time --end-date-time
Optional: --top
Example: ask-marcel list-calendar-view-delta --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z' --top 50
Graph: GET /me/calendarView/delta()?startDateTime={start-date-time}&endDateTime={end-date-time}

## list-calendars
List the calendars in the signed-in user’s mailbox (default + secondary calendars + shared calendars).
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-calendars
Graph: GET /me/calendars

## list-group-calendar-view
Return a date-windowed calendar view from a unified (Microsoft 365) group's calendar. Recurring events are expanded into individual occurrences across the window. Only Microsoft 365 groups have a calendar — security and distribution groups return `MailboxNotEnabledForRESTAPI`.
Required: --group-id --start-date-time --end-date-time
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-calendar-view --group-id 'a1b2c3d4-...' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'
Graph: GET /groups/{group-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}

## list-group-events
List events from a unified (Microsoft 365) group's calendar. Only Microsoft 365 groups have a calendar — security and distribution groups return an empty `value[]` or 404.
Required: --group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-events --group-id 'a1b2c3d4-...'
Graph: GET /groups/{group-id}/events

## list-room-lists
List room lists — usually one per building. Use these to scope a room search by location: a roomList groups the rooms in one office, then `/places/{roomList}/rooms` lists just those rooms. Pass `--top N` to limit the response on large tenants.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-room-lists
Graph: GET /places/microsoft.graph.roomList

## list-rooms
List bookable meeting rooms in the tenant. Each `room` has `displayName`, `emailAddress`, `capacity`, `building`, `floorNumber`, and `isWheelChairAccessible`. Use the `emailAddress` as a meeting `attendee` for room booking. Pass `--top 5` to limit the response — large tenants return tens of KB by de
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-rooms
Graph: GET /places/microsoft.graph.room

## list-shared-calendar-events
List events from another user's primary calendar (shared / delegated access). 403 without `Calendars.Read.Shared`.
Required: --user-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-shared-calendar-events --user-id 'colleague@contoso.com'
Graph: GET /users/{user-id}/calendar/events

## list-shared-calendar-view
Return a date-windowed calendar view from another user's primary calendar (shared / delegated access). Recurrences expanded into individual occurrences.
Required: --user-id --start-date-time --end-date-time
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-shared-calendar-view --user-id 'colleague@contoso.com' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'
Graph: GET /users/{user-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}

## list-specific-calendar-events
List the events in a specific calendar (does not expand recurrences). `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar (`/me/calendar/events`); any other value goes to `/me/calendars/{id}/events` and must be a real calendar ID.
Required: --calendar-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-specific-calendar-events --calendar-id 'primary'
Graph: GET /me/calendars/{calendar-id}/events

## list-specific-calendar-view
List the events in a specific calendar with recurrence expanded into individual occurrences in a date range. Both ISO date-time params are required by Graph. `--calendar-id primary` (or `default`) routes to the signed-in user’s default calendar.
Required: --calendar-id --start-date-time --end-date-time
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-specific-calendar-view --calendar-id 'primary' --start-date-time '2026-04-01T00:00:00Z' --end-date-time '2026-05-01T00:00:00Z'
Graph: GET /me/calendars/{calendar-id}/calendarView?startDateTime={start-date-time}&endDateTime={end-date-time}

## resolve-calendar-link
Parse a Microsoft Outlook calendar item link (the URL emitted by the "Copy link" / share action on a calendar event) into its `eventId`. Pure transformation — no Graph call. Pipe the result into `get-calendar-event` to fetch the event body. For Outlook mail message links use `resolve-mail-link` inst
Required: --url
Example: ask-marcel resolve-calendar-link --url 'https://outlook.office.com/calendar/item/AAMkAGI2THVS...'
Graph: GET {url}
