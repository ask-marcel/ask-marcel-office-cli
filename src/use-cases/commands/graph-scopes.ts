/**
 * Central Microsoft Graph permission scope map (audit round-8 Wave C).
 *
 * Source of truth: Microsoft Graph permissions reference
 * (https://learn.microsoft.com/en-us/graph/permissions-reference). Each
 * command's `graphDocsUrl` Learn page documents the required scope(s).
 *
 * Convention: list the LEAST-PRIVILEGED scope Microsoft documents for the
 * delegated permissions path. When multiple alternatives exist (e.g.
 * `Mail.Read` OR `Mail.ReadBasic` OR `Mail.ReadWrite`) we pick the most
 * read-only choice. The basic Teams web client token grants ~30 scopes
 * (`User.Read`, `Mail.Read`, `Calendars.Read`, `Files.Read`, `Tasks.Read`,
 * `Notes.Read.All`, `Sites.Read.All`, `Group.Read.All`, `Team.ReadBasic.All`,
 * `Channel.ReadBasic.All`, `People.Read`, `MailboxSettings.Read`, etc.). The
 * M365ChatClient elevated identity adds `Chat.ReadBasic` and ODSP allow-list
 * access (covered by `needsElevatedToken: true` per command).
 *
 * Surface: this map is read by `docs.ts toEntry` at manifest-render time
 * and merged into each command's `scopesRequired`. Per-command inline
 * `meta.scopesRequired` takes precedence (so command files can override).
 * An LLM uses this with `scopes-check` to predict 403s pre-flight.
 *
 * Lifecycle / meta commands intentionally absent (empty array would imply
 * "no scope needed" which is true for them but also true of `update`,
 * `docs`, `help-json` — none of which call Graph).
 */

const GRAPH_SCOPES_BY_COMMAND: Readonly<Record<string, ReadonlyArray<string>>> = {
  // — Drive / Files —
  'list-drives': ['Files.Read'],
  'get-drive-root-item': ['Files.Read'],
  'get-drive-special-folder': ['Files.Read'],
  'list-folder-files': ['Files.Read'],
  'get-drive-item': ['Files.Read'],
  'download-onedrive-file-content': ['Files.Read'],
  'list-drive-item-permissions': ['Files.Read'],
  'list-drive-item-versions': ['Files.Read'],
  'list-drive-item-thumbnails': ['Files.Read'],
  'search-onedrive-files': ['Files.Read'],
  'search-my-documents': ['Files.Read'],
  'get-drive-delta': ['Files.Read'],
  'get-drive-root-delta': ['Files.Read'],
  'list-recent-files': ['Files.Read'],
  'list-followed-drive-items': ['Files.Read'],
  'list-shared-with-me': ['Files.Read'],
  'get-drive-item-analytics': ['Files.Read.All', 'Sites.Read.All'],
  'get-drive-item-created-by-user': ['Files.Read'],
  'get-drive-item-last-modified-by-user': ['Files.Read'],
  'get-drive-item-list-item': ['Files.Read', 'Sites.Read.All'],
  'download-drive-item-as-pdf': ['Files.Read'],
  'download-drive-item-as-markdown': ['Files.Read'],
  'extract-drive-item-images': ['Files.Read'],
  'download-drive-item-version': ['Files.Read'],
  'list-trending-insights': ['Sites.Read.All'],
  'list-recently-used-insights': ['Sites.Read.All'],
  'list-shared-insights': ['Sites.Read.All'],

  // — Excel —
  'list-excel-worksheets': ['Files.Read'],
  'list-excel-tables': ['Files.Read'],
  'list-excel-table-rows': ['Files.Read'],
  'get-excel-table': ['Files.Read'],
  'get-excel-range': ['Files.Read'],
  'get-excel-used-range': ['Files.Read'],
  'list-excel-defined-names': ['Files.Read'],
  'list-excel-worksheet-charts': ['Files.Read'],
  'list-excel-worksheet-pivot-tables': ['Files.Read'],
  'list-excel-comments': ['Files.Read'],

  // — SharePoint —
  'search-sharepoint-sites-by-name': ['Sites.Read.All'],
  'get-sharepoint-site': ['Sites.Read.All'],
  'get-sharepoint-site-by-path': ['Sites.Read.All'],
  'get-sharepoint-site-drive-by-id': ['Sites.Read.All', 'Files.Read'],
  'list-sharepoint-site-drives': ['Sites.Read.All', 'Files.Read'],
  'list-sharepoint-site-lists': ['Sites.Read.All'],
  'get-sharepoint-site-list': ['Sites.Read.All'],
  'list-sharepoint-site-list-items': ['Sites.Read.All'],
  'get-sharepoint-site-list-item': ['Sites.Read.All'],
  'list-sharepoint-list-columns': ['Sites.Read.All'],
  'list-site-columns': ['Sites.Read.All'],
  'get-sharepoint-list-column': ['Sites.Read.All'],
  'list-sharepoint-list-item-versions': ['Sites.Read.All'],
  'list-site-content-types': ['Sites.Read.All'],
  'list-sharepoint-site-pages': ['Sites.Read.All'],
  'get-site-analytics': ['Sites.Read.All'],

  // — Mail / Outlook —
  'list-mail-messages': ['Mail.Read'],
  'list-mail-folders': ['Mail.Read'],
  'list-mail-folders-delta': ['Mail.Read'],
  'list-mail-child-folders': ['Mail.Read'],
  'list-mail-folder-messages': ['Mail.Read'],
  'list-mail-folder-messages-delta': ['Mail.Read'],
  'get-mail-message': ['Mail.Read'],
  'get-mail-message-mime': ['Mail.Read'],
  'list-mail-attachments': ['Mail.Read'],
  'get-mail-attachment': ['Mail.Read'],
  'list-mail-rules': ['MailboxSettings.Read'],
  'get-mail-rule': ['MailboxSettings.Read'],
  'get-mailbox-settings': ['MailboxSettings.Read'],
  'search-mail-messages': ['Mail.Read'],
  'list-outlook-categories': ['MailboxSettings.Read'],
  'list-focused-inbox-overrides': ['MailboxSettings.Read'],
  'convert-mail-to-markdown': ['Mail.Read'],
  'convert-mail-attachment-to-pdf': ['Mail.Read', 'Files.Read'],
  'convert-mail-attachment-to-markdown': ['Mail.Read'],
  'extract-mail-attachment-images': ['Mail.Read'],
  'extract-sharepoint-links-in-mail': ['Mail.Read'],
  'list-shared-mailbox-messages': ['Mail.Read.Shared'],
  'list-shared-mailbox-folder-messages': ['Mail.Read.Shared'],
  'get-shared-mailbox-message': ['Mail.Read.Shared'],
  'list-conversation-messages': ['Mail.Read'],

  // — Calendar —
  'list-calendar-events': ['Calendars.Read'],
  'get-calendar-event': ['Calendars.Read'],
  'list-specific-calendar-events': ['Calendars.Read'],
  'get-specific-calendar-event': ['Calendars.Read'],
  'list-calendar-view': ['Calendars.Read'],
  'list-specific-calendar-view': ['Calendars.Read'],
  'list-calendar-event-instances': ['Calendars.Read'],
  'list-calendar-events-delta': ['Calendars.Read'],
  'list-calendar-view-delta': ['Calendars.Read'],
  'list-calendars': ['Calendars.Read'],
  'list-calendar-groups': ['Calendars.Read'],
  'list-calendar-group-calendars': ['Calendars.Read'],
  'get-my-calendar': ['Calendars.Read'],
  'list-shared-calendar-events': ['Calendars.Read.Shared'],
  'list-shared-calendar-view': ['Calendars.Read.Shared'],
  'list-group-events': ['Group.Read.All', 'Calendars.Read'],
  'list-group-calendar-view': ['Group.Read.All', 'Calendars.Read'],

  // — Groups / Group-mailbox conversations —
  'get-group': ['GroupMember.Read.All'],
  'list-groups': ['GroupMember.Read.All'],
  'list-group-members': ['GroupMember.Read.All'],
  'list-group-owners': ['GroupMember.Read.All'],
  'list-group-conversations': ['Group.Read.All'],
  'list-group-threads': ['Group.Read.All'],

  // — Notes (OneNote) —
  'list-onenote-notebooks': ['Notes.Read'],
  'list-onenote-notebook-sections': ['Notes.Read'],
  'list-all-onenote-sections': ['Notes.Read'],
  'list-onenote-section-pages': ['Notes.Read'],
  'get-onenote-page-content': ['Notes.Read'],
  'get-onenote-page-as-markdown': ['Notes.Read'],
  'search-onenote-pages': ['Notes.Read'],
  'list-sharepoint-site-onenote-notebooks': ['Notes.Read.All', 'Sites.Read.All'],
  'list-sharepoint-site-onenote-notebook-sections': ['Notes.Read.All', 'Sites.Read.All'],
  'list-sharepoint-site-onenote-section-pages': ['Notes.Read.All', 'Sites.Read.All'],
  'get-sharepoint-site-onenote-page-content': ['Notes.Read.All', 'Sites.Read.All'],

  // — Tasks (To Do + Planner) —
  'list-todo-task-lists': ['Tasks.Read'],
  'list-todo-tasks': ['Tasks.Read'],
  'list-todo-tasks-delta': ['Tasks.Read'],
  'list-incomplete-todo-tasks': ['Tasks.Read'],
  'get-todo-task': ['Tasks.Read'],
  'list-todo-linked-resources': ['Tasks.Read'],
  'list-planner-plans': ['Tasks.Read'],
  'list-plan-buckets': ['Tasks.Read'],
  'list-plan-tasks': ['Tasks.Read'],
  'list-planner-tasks': ['Tasks.Read'],
  'list-incomplete-planner-tasks': ['Tasks.Read'],
  'get-planner-plan': ['Tasks.Read'],
  'get-planner-task': ['Tasks.Read'],
  'get-planner-task-details': ['Tasks.Read'],
  'get-planner-bucket': ['Tasks.Read'],

  // — User / People / Org —
  'get-current-user': ['User.Read'],
  'get-my-manager': ['User.Read.All'],
  'get-user-manager': ['User.Read.All'],
  'get-my-profile-photo': ['User.Read'],
  'list-my-direct-reports': ['User.Read.All'],
  'list-user-direct-reports': ['User.Read.All'],
  'list-my-memberships': ['User.Read'],
  'list-my-transitive-memberships': ['User.Read'],
  'list-relevant-people': ['People.Read'],
  'get-organization': ['User.Read'],
  'list-sensitivity-labels': ['InformationProtectionPolicy.Read'],

  // — Teams —
  'list-joined-teams': ['Team.ReadBasic.All'],
  'get-team': ['Team.ReadBasic.All'],
  'list-team-channels': ['Channel.ReadBasic.All'],
  'get-team-channel': ['Channel.ReadBasic.All'],
  'get-team-primary-channel': ['Channel.ReadBasic.All'],
  'list-team-installed-apps': ['TeamsAppInstallation.ReadForTeam'],
  'get-channel-files-folder': ['Channel.ReadBasic.All', 'Files.Read'],

  // — Chats (require M365ChatClient elevated token; see needsElevatedToken) —
  'list-chats': ['Chat.ReadBasic'],
  'get-chat': ['Chat.ReadBasic'],
  'list-chat-members': ['ChatMember.Read'],
  // chatsvcagg-tier commands — auth gates server-side on the captured Teams
  // client identity, not on Graph delegated scopes. Empty array means "no
  // Graph scope required" (the bearer audience is chatsvcagg.teams.microsoft.com,
  // not graph.microsoft.com).
  'list-teams-chats-with-messages': [],
  'list-teams-chat-messages': [],
  'list-teams-chat-history': [],
  'get-teams-chat-message': [],
  'resolve-teams-link': [],
  'find-chats-with-user': [],

  // — Places (rooms) —
  'list-rooms': ['Place.Read.All'],
  'list-room-lists': ['Place.Read.All'],

  // — Meta —
  'my-quick-context': ['User.Read', 'Files.Read', 'Mail.Read', 'Tasks.Read', 'Calendars.Read', 'Notes.Read', 'Team.ReadBasic.All'],
  // microsoft-search-query, next-page: scopes inherited from the underlying
  // entity types / cursor target — left unspecified to avoid misleading.
  // scopes-check: cached-token introspection, no Graph call.
};

const lookupScopes = (commandName: string): ReadonlyArray<string> | undefined => GRAPH_SCOPES_BY_COMMAND[commandName];

export { GRAPH_SCOPES_BY_COMMAND, lookupScopes };
