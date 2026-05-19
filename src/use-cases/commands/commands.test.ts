import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';
import type { FetchFn, GraphError } from '../../infra/graph-client.ts';
import { createGraphClient } from '../../infra/graph-client.ts';
import { buildSampleDocx, buildSampleXlsx } from '../../test-helpers/office-fixtures.ts';
import { renderSingleCommand } from './docs.ts';
import { commands as cmdRegistry } from './index.ts';
import * as downloadDriveItemAsMarkdown from './download-drive-item-as-markdown.ts';
import * as downloadDriveItemAsPdf from './download-drive-item-as-pdf.ts';
import * as downloadDriveItemVersionAsMarkdown from './download-drive-item-version-as-markdown.ts';
import * as downloadDriveItemVersionAsPdf from './download-drive-item-version-as-pdf.ts';
import * as downloadDriveItemVersionContent from './download-drive-item-version-content.ts';
import * as downloadOnedriveFileContent from './download-onedrive-file-content.ts';
import * as getCalendarEvent from './get-calendar-event.ts';
import * as getCalendarView from './get-calendar-view.ts';
import * as getCurrentUser from './get-current-user.ts';
import * as getDriveDelta from './get-drive-delta.ts';
import * as getDriveItem from './get-drive-item.ts';
import * as getDriveRootItem from './get-drive-root-item.ts';
import * as getExcelRange from './get-excel-range.ts';
import * as getExcelTable from './get-excel-table.ts';
import * as getMailAttachment from './get-mail-attachment.ts';
import * as getMailMessage from './get-mail-message.ts';
import * as getMailboxSettings from './get-mailbox-settings.ts';
import * as getMyProfilePhoto from './get-my-profile-photo.ts';
import * as getOnenotePageAsMarkdown from './get-onenote-page-as-markdown.ts';
import * as getOnenotePageContent from './get-onenote-page-content.ts';
import * as getPlannerBucket from './get-planner-bucket.ts';
import * as getPlannerPlan from './get-planner-plan.ts';
import * as getPlannerTaskDetails from './get-planner-task-details.ts';
import * as getPlannerTask from './get-planner-task.ts';
import * as getSharepointSiteByPath from './get-sharepoint-site-by-path.ts';
import * as getSharepointSiteDriveById from './get-sharepoint-site-drive-by-id.ts';
import * as getSharepointSiteListItem from './get-sharepoint-site-list-item.ts';
import * as getSharepointSiteList from './get-sharepoint-site-list.ts';
import * as getSharepointSite from './get-sharepoint-site.ts';
import * as getSpecificCalendarEvent from './get-specific-calendar-event.ts';
import * as getSpecificCalendarView from './get-specific-calendar-view.ts';
import * as getTeamChannel from './get-team-channel.ts';
import * as getTeam from './get-team.ts';
import * as getTodoTask from './get-todo-task.ts';
import * as listAllOnenoteSections from './list-all-onenote-sections.ts';
import * as listCalendarEventInstances from './list-calendar-event-instances.ts';
import * as listCalendarEventsDelta from './list-calendar-events-delta.ts';
import * as listCalendarEvents from './list-calendar-events.ts';
import * as listCalendarViewDelta from './list-calendar-view-delta.ts';
import * as listCalendars from './list-calendars.ts';
import * as listChatMembers from './list-chat-members.ts';
import * as listDriveItemPermissions from './list-drive-item-permissions.ts';
import * as listDriveItemVersions from './list-drive-item-versions.ts';
import * as listDrives from './list-drives.ts';
import * as listExcelTableRows from './list-excel-table-rows.ts';
import * as listExcelTables from './list-excel-tables.ts';
import * as listExcelWorksheets from './list-excel-worksheets.ts';
import * as listFolderFiles from './list-folder-files.ts';
import * as listIncompletePlannerTasks from './list-incomplete-planner-tasks.ts';
import * as listIncompleteTodoTasks from './list-incomplete-todo-tasks.ts';
import * as listJoinedTeams from './list-joined-teams.ts';
import * as listMailAttachments from './list-mail-attachments.ts';
import * as listMailChildFolders from './list-mail-child-folders.ts';
import * as listMailFolderMessages from './list-mail-folder-messages.ts';
import * as listMailFolders from './list-mail-folders.ts';
import * as listMailMessages from './list-mail-messages.ts';
import * as listMailRules from './list-mail-rules.ts';
import * as listOnenoteNotebookSections from './list-onenote-notebook-sections.ts';
import * as listOnenoteNotebooks from './list-onenote-notebooks.ts';
import * as listOnenoteSectionPages from './list-onenote-section-pages.ts';
import * as listPlanBuckets from './list-plan-buckets.ts';
import * as listPlanTasks from './list-plan-tasks.ts';
import * as listPlannerPlans from './list-planner-plans.ts';
import * as listPlannerTasks from './list-planner-tasks.ts';
import * as listSharepointSiteDrives from './list-sharepoint-site-drives.ts';
import * as listSharepointSiteListItems from './list-sharepoint-site-list-items.ts';
import * as listSharepointSiteLists from './list-sharepoint-site-lists.ts';
import * as listSpecificCalendarEvents from './list-specific-calendar-events.ts';
import * as listTeamChannels from './list-team-channels.ts';
import * as listTodoLinkedResources from './list-todo-linked-resources.ts';
import * as listTodoTaskLists from './list-todo-task-lists.ts';
import * as listTodoTasks from './list-todo-tasks.ts';
import * as nextPage from './next-page.ts';
import * as searchMailMessages from './search-mail-messages.ts';
import * as searchMyDocuments from './search-my-documents.ts';
import * as searchOnedriveFiles from './search-onedrive-files.ts';
import * as searchOnenotePages from './search-onenote-pages.ts';
import * as searchSharepointSitesByName from './search-sharepoint-sites-by-name.ts';
import * as extractSharepointLinksInMail from './extract-sharepoint-links-in-mail.ts';
import * as convertMailAttachmentToMarkdown from './convert-mail-attachment-to-markdown.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as convertMailToMarkdown from './convert-mail-to-markdown.ts';
import * as listChats from './list-chats.ts';
import * as getChat from './get-chat.ts';
import * as listMyDirectReports from './list-my-direct-reports.ts';
import * as listUserDirectReports from './list-user-direct-reports.ts';
import * as listRecentFiles from './list-recent-files.ts';
import * as listSharedWithMe from './list-shared-with-me.ts';
import * as listRecentlyUsedInsights from './list-recently-used-insights.ts';
import * as listSharedInsights from './list-shared-insights.ts';
import * as getOrganization from './get-organization.ts';
import * as listMailFoldersDelta from './list-mail-folders-delta.ts';
import * as getChannelFilesFolder from './get-channel-files-folder.ts';
import * as getDriveItemListItem from './get-drive-item-list-item.ts';
import * as getDriveItemAnalytics from './get-drive-item-analytics.ts';
import * as listTeamInstalledApps from './list-team-installed-apps.ts';
import * as listCalendarGroups from './list-calendar-groups.ts';
import * as listCalendarGroupCalendars from './list-calendar-group-calendars.ts';
import * as getMyCalendar from './get-my-calendar.ts';
import * as listSiteColumns from './list-site-columns.ts';
import * as listSiteContentTypes from './list-site-content-types.ts';
import * as listSharepointSitePages from './list-sharepoint-site-pages.ts';
import * as listExcelDefinedNames from './list-excel-defined-names.ts';
import * as listExcelWorksheetCharts from './list-excel-worksheet-charts.ts';
import * as microsoftSearchQuery from './microsoft-search-query.ts';
import * as getDriveSpecialFolder from './get-drive-special-folder.ts';
import * as getDriveRootDelta from './get-drive-root-delta.ts';
import * as listFollowedDriveItems from './list-followed-drive-items.ts';
import * as getDriveItemCreatedByUser from './get-drive-item-created-by-user.ts';
import * as getDriveItemLastModifiedByUser from './get-drive-item-last-modified-by-user.ts';
import * as getSiteAnalytics from './get-site-analytics.ts';
import * as listSharepointListItemVersions from './list-sharepoint-list-item-versions.ts';
import * as getMailRule from './get-mail-rule.ts';
import * as listExcelComments from './list-excel-comments.ts';
import * as listExcelWorksheetPivotTables from './list-excel-worksheet-pivot-tables.ts';
import * as listSensitivityLabels from './list-sensitivity-labels.ts';
import * as listMyTransitiveMemberships from './list-my-transitive-memberships.ts';
import * as getTeamPrimaryChannel from './get-team-primary-channel.ts';
import * as listTodoTasksDelta from './list-todo-tasks-delta.ts';
import * as listMyMemberships from './list-my-memberships.ts';
import * as getMyManager from './get-my-manager.ts';
import * as getUserManager from './get-user-manager.ts';
import * as listRelevantPeople from './list-relevant-people.ts';
import * as listGroups from './list-groups.ts';
import * as getGroup from './get-group.ts';
import * as listGroupMembers from './list-group-members.ts';
import * as listGroupOwners from './list-group-owners.ts';
import * as listGroupEvents from './list-group-events.ts';
import * as getGroupCalendarView from './get-group-calendar-view.ts';
import * as listGroupConversations from './list-group-conversations.ts';
import * as listGroupThreads from './list-group-threads.ts';
import * as getMailMessageMime from './get-mail-message-mime.ts';
import * as listMailFolderMessagesDelta from './list-mail-folder-messages-delta.ts';
import * as listSharedMailboxMessages from './list-shared-mailbox-messages.ts';
import * as listSharedMailboxFolderMessages from './list-shared-mailbox-folder-messages.ts';
import * as getSharedMailboxMessage from './get-shared-mailbox-message.ts';
import * as listConversationMessages from './list-conversation-messages.ts';
import * as listFocusedInboxOverrides from './list-focused-inbox-overrides.ts';
import * as listOutlookCategories from './list-outlook-categories.ts';
import * as listSharedCalendarEvents from './list-shared-calendar-events.ts';
import * as getSharedCalendarView from './get-shared-calendar-view.ts';
import * as listSharepointListColumns from './list-sharepoint-list-columns.ts';
import * as getSharepointListColumn from './get-sharepoint-list-column.ts';
import * as listSharepointSiteOnenoteNotebooks from './list-sharepoint-site-onenote-notebooks.ts';
import * as listSharepointSiteOnenoteNotebookSections from './list-sharepoint-site-onenote-notebook-sections.ts';
import * as listSharepointSiteOnenoteSectionPages from './list-sharepoint-site-onenote-section-pages.ts';
import * as getSharepointSiteOnenotePageContent from './get-sharepoint-site-onenote-page-content.ts';
import * as listDriveItemThumbnails from './list-drive-item-thumbnails.ts';
import * as getExcelUsedRange from './get-excel-used-range.ts';
import * as listRooms from './list-rooms.ts';
import * as listRoomLists from './list-room-lists.ts';
import * as listTrendingInsights from './list-trending-insights.ts';

const cmdMap: Record<string, { execute: typeof listDrives.execute }> = {
  'list-drives': listDrives,
  'get-drive-root-item': getDriveRootItem,
  'list-folder-files': listFolderFiles,
  'download-onedrive-file-content': downloadOnedriveFileContent,
  'get-drive-item': getDriveItem,
  'list-drive-item-permissions': listDriveItemPermissions,
  'list-drive-item-versions': listDriveItemVersions,
  'download-drive-item-version-content': downloadDriveItemVersionContent,
  'download-drive-item-as-pdf': downloadDriveItemAsPdf,
  'download-drive-item-as-markdown': downloadDriveItemAsMarkdown,
  'download-drive-item-version-as-pdf': downloadDriveItemVersionAsPdf,
  'download-drive-item-version-as-markdown': downloadDriveItemVersionAsMarkdown,
  'search-onedrive-files': searchOnedriveFiles,
  'search-my-documents': searchMyDocuments,
  'get-excel-range': getExcelRange,
  'list-excel-worksheets': listExcelWorksheets,
  'list-excel-tables': listExcelTables,
  'get-excel-table': getExcelTable,
  'list-excel-table-rows': listExcelTableRows,
  'get-drive-delta': getDriveDelta,
  'search-sharepoint-sites-by-name': searchSharepointSitesByName,
  'get-sharepoint-site': getSharepointSite,
  'list-sharepoint-site-drives': listSharepointSiteDrives,
  'get-sharepoint-site-drive-by-id': getSharepointSiteDriveById,
  'list-sharepoint-site-lists': listSharepointSiteLists,
  'get-sharepoint-site-list': getSharepointSiteList,
  'list-sharepoint-site-list-items': listSharepointSiteListItems,
  'get-sharepoint-site-list-item': getSharepointSiteListItem,
  'get-sharepoint-site-by-path': getSharepointSiteByPath,
  'list-todo-task-lists': listTodoTaskLists,
  'list-todo-tasks': listTodoTasks,
  'list-incomplete-todo-tasks': listIncompleteTodoTasks,
  'get-todo-task': getTodoTask,
  'list-todo-linked-resources': listTodoLinkedResources,
  'list-planner-plans': listPlannerPlans,
  'list-planner-tasks': listPlannerTasks,
  'list-incomplete-planner-tasks': listIncompletePlannerTasks,
  'get-planner-plan': getPlannerPlan,
  'list-plan-tasks': listPlanTasks,
  'get-planner-task': getPlannerTask,
  'get-planner-task-details': getPlannerTaskDetails,
  'list-plan-buckets': listPlanBuckets,
  'get-planner-bucket': getPlannerBucket,
  'list-mail-messages': listMailMessages,
  'list-mail-folders': listMailFolders,
  'list-mail-child-folders': listMailChildFolders,
  'list-mail-folder-messages': listMailFolderMessages,
  'get-mail-message': getMailMessage,
  'list-mail-attachments': listMailAttachments,
  'get-mail-attachment': getMailAttachment,
  'list-mail-rules': listMailRules,
  'get-mailbox-settings': getMailboxSettings,
  'search-mail-messages': searchMailMessages,
  'extract-sharepoint-links-in-mail': extractSharepointLinksInMail,
  'convert-mail-to-markdown': convertMailToMarkdown,
  'convert-mail-attachment-to-pdf': convertMailAttachmentToPdf,
  'convert-mail-attachment-to-markdown': convertMailAttachmentToMarkdown,
  'list-onenote-notebooks': listOnenoteNotebooks,
  'list-onenote-notebook-sections': listOnenoteNotebookSections,
  'list-all-onenote-sections': listAllOnenoteSections,
  'list-onenote-section-pages': listOnenoteSectionPages,
  'get-onenote-page-content': getOnenotePageContent,
  'get-onenote-page-as-markdown': getOnenotePageAsMarkdown,
  'search-onenote-pages': searchOnenotePages,
  'get-current-user': getCurrentUser,
  'get-my-profile-photo': getMyProfilePhoto,
  'list-calendar-events': listCalendarEvents,
  'get-calendar-event': getCalendarEvent,
  'list-specific-calendar-events': listSpecificCalendarEvents,
  'get-specific-calendar-event': getSpecificCalendarEvent,
  'list-calendar-view': getCalendarView,
  'list-specific-calendar-view': getSpecificCalendarView,
  'list-calendar-event-instances': listCalendarEventInstances,
  'list-calendars': listCalendars,
  'list-calendar-events-delta': listCalendarEventsDelta,
  'list-calendar-view-delta': listCalendarViewDelta,
  'list-chat-members': listChatMembers,
  'list-joined-teams': listJoinedTeams,
  'get-team': getTeam,
  'list-team-channels': listTeamChannels,
  'get-team-channel': getTeamChannel,
  'list-chats': listChats,
  'get-chat': getChat,
  'list-my-direct-reports': listMyDirectReports,
  'list-user-direct-reports': listUserDirectReports,
  'list-my-memberships': listMyMemberships,
  'get-my-manager': getMyManager,
  'get-user-manager': getUserManager,
  'list-relevant-people': listRelevantPeople,
  'list-groups': listGroups,
  'get-group': getGroup,
  'list-group-members': listGroupMembers,
  'list-group-owners': listGroupOwners,
  'list-group-events': listGroupEvents,
  'list-group-calendar-view': getGroupCalendarView,
  'list-group-conversations': listGroupConversations,
  'list-group-threads': listGroupThreads,
  'get-mail-message-mime': getMailMessageMime,
  'list-mail-folder-messages-delta': listMailFolderMessagesDelta,
  'list-shared-mailbox-messages': listSharedMailboxMessages,
  'list-shared-mailbox-folder-messages': listSharedMailboxFolderMessages,
  'get-shared-mailbox-message': getSharedMailboxMessage,
  'list-conversation-messages': listConversationMessages,
  'list-focused-inbox-overrides': listFocusedInboxOverrides,
  'list-outlook-categories': listOutlookCategories,
  'list-shared-calendar-events': listSharedCalendarEvents,
  'list-shared-calendar-view': getSharedCalendarView,
  'list-sharepoint-list-columns': listSharepointListColumns,
  'get-sharepoint-list-column': getSharepointListColumn,
  'list-sharepoint-site-onenote-notebooks': listSharepointSiteOnenoteNotebooks,
  'list-sharepoint-site-onenote-notebook-sections': listSharepointSiteOnenoteNotebookSections,
  'list-sharepoint-site-onenote-section-pages': listSharepointSiteOnenoteSectionPages,
  'get-sharepoint-site-onenote-page-content': getSharepointSiteOnenotePageContent,
  'list-drive-item-thumbnails': listDriveItemThumbnails,
  'get-excel-used-range': getExcelUsedRange,
  'list-rooms': listRooms,
  'list-room-lists': listRoomLists,
  'list-trending-insights': listTrendingInsights,
  'list-recent-files': listRecentFiles,
  'list-shared-with-me': listSharedWithMe,
  'list-recently-used-insights': listRecentlyUsedInsights,
  'list-shared-insights': listSharedInsights,
  'get-organization': getOrganization,
  'list-mail-folders-delta': listMailFoldersDelta,
  'get-channel-files-folder': getChannelFilesFolder,
  'get-drive-item-list-item': getDriveItemListItem,
  'get-drive-item-analytics': getDriveItemAnalytics,
  'list-team-installed-apps': listTeamInstalledApps,
  'list-calendar-groups': listCalendarGroups,
  'list-calendar-group-calendars': listCalendarGroupCalendars,
  'get-my-calendar': getMyCalendar,
  'list-site-columns': listSiteColumns,
  'list-site-content-types': listSiteContentTypes,
  'list-sharepoint-site-pages': listSharepointSitePages,
  'list-excel-defined-names': listExcelDefinedNames,
  'list-excel-worksheet-charts': listExcelWorksheetCharts,
  'microsoft-search-query': microsoftSearchQuery,
  'get-drive-special-folder': getDriveSpecialFolder,
  'get-drive-root-delta': getDriveRootDelta,
  'list-followed-drive-items': listFollowedDriveItems,
  'get-drive-item-created-by-user': getDriveItemCreatedByUser,
  'get-drive-item-last-modified-by-user': getDriveItemLastModifiedByUser,
  'get-site-analytics': getSiteAnalytics,
  'list-sharepoint-list-item-versions': listSharepointListItemVersions,
  'get-mail-rule': getMailRule,
  'list-excel-comments': listExcelComments,
  'list-excel-worksheet-pivot-tables': listExcelWorksheetPivotTables,
  'list-sensitivity-labels': listSensitivityLabels,
  'list-my-transitive-memberships': listMyTransitiveMemberships,
  'get-team-primary-channel': getTeamPrimaryChannel,
  'list-todo-tasks-delta': listTodoTasksDelta,
  'next-page': nextPage,
};

const fakeAuth = (): AuthManager => ({
  getAccessToken: async () => ok(accessTokenUnsafe('test-token')),
  getElevatedAccessToken: async () => ok(accessTokenUnsafe('test-elevated-token')),
  logout: async () => ok(undefined),
  getLastElevatedOutcome: () => null,
});

type FakeFetch = ((url: string, init?: RequestInit) => Promise<Response>) & { lastUrl: string | null; lastBody: string | null };

const fakeFetch = (body: unknown): FakeFetch => {
  let lastUrl: string | null = null;
  let lastBody: string | null = null;
  const fn = async (url: string, init?: RequestInit): Promise<Response> => {
    lastUrl = url;
    lastBody = typeof init?.body === 'string' ? init.body : null;
    return new Response(JSON.stringify(body), { headers: { 'content-type': 'application/json' } });
  };
  Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
  Object.defineProperty(fn, 'lastBody', { get: () => lastBody });
  return fn as FakeFetch;
};

const callCommand = async (name: string, params: Record<string, string>, responseBody: unknown): Promise<Result<unknown, GraphError>> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch(responseBody);
  const graph = createGraphClient(fakeAuth(), fetchFn);
  return cmd.execute(graph, params);
};

const capturedUrl = async (name: string, params: Record<string, string>): Promise<string> => {
  const cmd = cmdMap[name];
  if (!cmd) throw new Error(`command not found: ${name}`);
  const fetchFn = fakeFetch({ ok: true });
  const graph = createGraphClient(fakeAuth(), fetchFn);
  await cmd.execute(graph, params);
  return fetchFn.lastUrl ?? '';
};

/**
 * Stateful fetch fake for multi-step commands (per the FIX #4 design):
 * each handler matches the FIRST entry whose `urlPrefix` is a prefix
 * of the request URL AND whose optional `method` matches. The matched
 * handler is then CONSUMED — subsequent calls fall through to the
 * next available handler. Unmatched calls throw a clear error rather
 * than silently mismatching. `response` may be a static `Response` or
 * a thunk so each call gets a fresh body / headers.
 */
type StagedHandler = {
  readonly urlPrefix: string;
  readonly method?: string;
  readonly response: Response | (() => Response);
};

const stagedFetch = (handlers: ReadonlyArray<StagedHandler>): ((url: string, init?: RequestInit) => Promise<Response>) => {
  const queue = handlers.map((h) => ({ ...h, consumed: false }));
  return async (url, init) => {
    const requestUrl = url.split('?')[0] ?? url;
    const requestUrlWithQuery = url;
    const method = init?.method ?? 'GET';
    const idx = queue.findIndex((h) => {
      if (h.consumed) return false;
      if (h.method !== undefined && h.method !== method) return false;
      const prefix = h.urlPrefix.split('?')[0] ?? h.urlPrefix;
      const queryNeeded = h.urlPrefix.includes('?');
      if (queryNeeded) return requestUrlWithQuery.startsWith(h.urlPrefix);
      return requestUrl.startsWith(prefix);
    });
    if (idx === -1) throw new Error(`stagedFetch: unexpected ${method} ${url}`);
    const handler = queue[idx];
    if (!handler) throw new Error(`stagedFetch: corrupt queue at ${idx}`);
    handler.consumed = true;
    return typeof handler.response === 'function' ? handler.response() : handler.response.clone();
  };
};

describe('commands', () => {
  it('list-drives returns drives', async () => {
    const result = await callCommand('list-drives', {}, { value: [{ id: 'd1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'd1' }] });
  });

  it('get-drive-root-item returns root item', async () => {
    const result = await callCommand('get-drive-root-item', { driveId: 'drive-1' }, { id: 'root', name: 'Root' });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ id: 'root', name: 'Root' });
  });

  it('list-folder-files returns children', async () => {
    const result = await callCommand('list-folder-files', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'file.txt' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'file.txt' }] });
  });

  it('get-drive-item returns metadata', async () => {
    const result = await callCommand('get-drive-item', { driveId: 'd1', itemId: 'i1' }, { name: 'doc.xlsx', size: 1024 });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ name: 'doc.xlsx', size: 1024 });
  });

  it('search-onedrive-files searches with query', async () => {
    const result = await callCommand('search-onedrive-files', { driveId: 'd1', query: 'report' }, { value: [{ name: 'report.xlsx' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'report.xlsx' }] });
  });

  it('search-my-documents searches the user’s default OneDrive', async () => {
    const result = await callCommand('search-my-documents', { query: 'budget' }, { value: [{ name: 'budget.xlsx' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'budget.xlsx' }] });
  });

  it('search-mail-messages searches the mailbox with $search', async () => {
    const result = await callCommand('search-mail-messages', { query: 'invoice' }, { value: [{ subject: 'invoice 042' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ subject: 'invoice 042' }] });
  });

  it('search-onenote-pages filters OneNote pages by title (Graph removed full-text ?search= from v1.0)', async () => {
    const result = await callCommand('search-onenote-pages', { titleSubstring: 'meeting notes' }, { value: [{ title: 'Meeting notes 2026-04-30' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ title: 'Meeting notes 2026-04-30' }] });
  });

  it('search-sharepoint-sites-by-name searches sites with the search query parameter', async () => {
    const result = await callCommand('search-sharepoint-sites-by-name', { query: 'marketing' }, { value: [{ displayName: 'Marketing site' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ displayName: 'Marketing site' }] });
  });

  it('list-incomplete-planner-tasks filters Planner tasks by percentComplete ne 100', async () => {
    const result = await callCommand('list-incomplete-planner-tasks', {}, { value: [{ id: 'pt1', percentComplete: 0 }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'pt1', percentComplete: 0 }] });
  });

  it('list-incomplete-todo-tasks filters To Do tasks by status ne completed within a list', async () => {
    const result = await callCommand('list-incomplete-todo-tasks', { todoTaskListId: 'tl1' }, { value: [{ id: 'tt1', status: 'inProgress' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'tt1', status: 'inProgress' }] });
  });

  it('next-page strips the Graph v1.0 prefix and GETs the rest of the supplied URL', async () => {
    const result = await callCommand('next-page', { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' }, { value: [{ subject: 'page 2' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ subject: 'page 2' }] });
  });

  it('next-page rejects URLs that do not start with the Graph v1.0 prefix as a validation_error Result', async () => {
    const cmd = cmdMap['next-page'];
    if (!cmd) throw new Error('next-page not registered');
    const fetchFn = fakeFetch({ ok: true });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { url: 'https://example.com/something' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('get-excel-range returns cell values', async () => {
    const result = await callCommand(
      'get-excel-range',
      { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1:B2' },
      {
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      }
    );
    expect(result.ok).toBe(true);
    if (result.ok)
      expect(result.value).toEqual({
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      });
  });

  it('get-excel-range rejects an --address spanning more than 100 000 cells', async () => {
    const result = await callCommand('get-excel-range', { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'ZZ999999:AAA1' }, {});
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).toContain('cells (cap: 100,000)');
      // Audit round-7 B2: formatZodError prepends `--address` itself, so the
      // superRefine message must not start with `--address ` too (or we get the
      // duplicate `--address --address spans …`).
      expect(result.error.message).not.toMatch(/--address --address/);
      expect(result.error.message).toMatch(/^--address spans/);
    }
  });

  it('get-excel-range accepts a single-cell --address without parsing', async () => {
    const result = await callCommand('get-excel-range', { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'B7' }, { values: [['Fendi']] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ values: [['Fendi']] });
  });

  it('list-excel-worksheets returns worksheets', async () => {
    const result = await callCommand('list-excel-worksheets', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'Sheet1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'Sheet1' }] });
  });

  it('list-excel-tables returns tables', async () => {
    const result = await callCommand('list-excel-tables', { driveId: 'd1', itemId: 'i1' }, { value: [{ name: 'Table1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ name: 'Table1' }] });
  });

  it('get-excel-table returns table details', async () => {
    const result = await callCommand('get-excel-table', { driveId: 'd1', itemId: 'i1', tableId: 't1' }, { name: 'Table1', showHeaders: true });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ name: 'Table1', showHeaders: true });
  });

  it('list-excel-table-rows returns rows', async () => {
    const result = await callCommand('list-excel-table-rows', { driveId: 'd1', itemId: 'i1', tableId: 't1' }, { value: [{ index: 0, values: ['a', 'b'] }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ index: 0, values: ['a', 'b'] }] });
  });

  it('get-drive-delta returns changes', async () => {
    const result = await callCommand('get-drive-delta', { driveId: 'd1', itemId: 'i1' }, { value: [{ id: 'new-file' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'new-file' }] });
  });

  it('list-drive-item-permissions returns permissions', async () => {
    const result = await callCommand('list-drive-item-permissions', { driveId: 'd1', itemId: 'i1' }, { value: [{ roles: ['read'] }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ roles: ['read'] }] });
  });

  it('list-drive-item-versions returns versions', async () => {
    const result = await callCommand('list-drive-item-versions', { driveId: 'd1', itemId: 'i1' }, { value: [{ id: 'v1' }] });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual({ value: [{ id: 'v1' }] });
  });

  it('download-onedrive-file-content inlines the bytes (no longer returns the bare downloadUrl envelope)', async () => {
    const result = await callCommand('download-onedrive-file-content', { driveId: 'd1', itemId: 'i1' }, { contentType: 'application/octet-stream', size: 5, base64: 'JVBERi0=' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/octet-stream');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-onedrive-file-content rejects a folder --item-id with a clear "this is a folder, use list-folder-files" hint instead of empty error (audit round-6 §1.1)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFolder', method: 'GET', response: Response.json({ name: 'Reports', folder: { childCount: 12 } }) },
    ]);
    const cmd = cmdMap['download-onedrive-file-content'];
    if (!cmd) throw new Error('download-onedrive-file-content not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iFolder' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(400);
      expect(result.error.message).toContain("item 'Reports' is a folder");
      expect(result.error.message).toContain('list-folder-files');
    }
  });

  it('download-drive-item-as-pdf rejects a folder --item-id with a clear "this is a folder, use list-folder-files" hint (audit round-6 §1.1)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFolder', method: 'GET', response: Response.json({ name: 'Reports', folder: { childCount: 12 } }) },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iFolder' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(400);
      expect(result.error.message).toContain("item 'Reports' is a folder");
      expect(result.error.message).toContain('list-folder-files');
    }
  });

  it('get-mail-attachment surfaces `base64` mirror of `contentBytes` for fileAttachments so --output-path can land on disk (audit round-6 §6)', async () => {
    const result = await callCommand(
      'get-mail-attachment',
      { messageId: 'm1', attachmentId: 'a1' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: 'JVBERi0=' }
    );
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentBytes: string; base64: string };
      expect(v.contentBytes).toBe('JVBERi0=');
      expect(v.base64).toBe('JVBERi0=');
    }
  });

  it('get-mail-attachment does NOT add a base64 mirror to itemAttachment / referenceAttachment (no raw bytes there)', async () => {
    const result = await callCommand(
      'get-mail-attachment',
      { messageId: 'm1', attachmentId: 'a1' },
      { '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://example.com/x' }
    );
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { base64?: string };
      expect(v.base64).toBeUndefined();
    }
  });

  it('list-todo-tasks passes through a NON-ParseUri error unchanged (only the known opaque case is rewritten)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { code: 'Unauthorized', message: 'bad token' } }), { status: 401, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['list-todo-tasks'];
    if (!cmd) throw new Error('list-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toBe('Unauthorized: bad token');
    }
  });

  it('list-todo-tasks rewrites the opaque RequestBroker--ParseUri error to a hint that names the workaround (audit round-6 §1.5)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks?$select=id%2Ctitle',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-todo-tasks'];
    if (!cmd) throw new Error('list-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', select: 'id,title' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('Graph rejected --select=id,title');
      expect(r.error.message).toContain('Drop `title` from --select');
    }
  });

  it('list-todo-tasks also rewrites RequestBroker--ParseUri when --orderby trips the same parser quirk (audit v1.0.0 Bug 3)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks?$orderby=title%20asc',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-todo-tasks'];
    if (!cmd) throw new Error('list-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', orderby: 'title asc' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('Graph rejected --orderby=title asc');
      expect(r.error.message).toContain('sorting on `title` is unsupported');
    }
  });

  it('list-todo-tasks passes through RequestBroker--ParseUri unchanged when neither --select nor --orderby is set (no false rewrite)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-todo-tasks'];
    if (!cmd) throw new Error('list-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('RequestBroker--ParseUri');
      expect(r.error.message).not.toContain('Drop `title`');
    }
  });

  it('list-calendar-event-instances rewrites the opaque ExpandSeries error to a seriesMaster hint (audit v1.0.0 Issue 9)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/calendar/events/e1/instances',
        method: 'GET',
        response: () =>
          new Response(
            JSON.stringify({
              error: {
                code: 'ErrorInvalidRequest',
                message: "Your request can't be completed. ExpandSeries can only be performed against a series.",
              },
            }),
            { status: 400, headers: { 'content-type': 'application/json' } }
          ),
      },
    ]);
    const cmd = cmdMap['list-calendar-event-instances'];
    if (!cmd) throw new Error('list-calendar-event-instances not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('not a recurring series');
      expect(r.error.message).toContain("type eq 'seriesMaster'");
    }
  });

  it('list-chat-members passes through non-`1: NotFound` Graph errors unchanged (no false rewrite)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/chats/19:abc@thread.v2/members',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'AccessDenied', message: 'User does not have permission.' } }), {
            status: 403,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-chat-members'];
    if (!cmd) throw new Error('list-chat-members not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { chatId: '19:abc@thread.v2' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('AccessDenied');
      expect(r.error.message).not.toContain('Microsoft Teams chat not found');
    }
  });

  it('list-chat-members rewrites the opaque `1: NotFound` error to a clear chat-id hint (audit round-7 B3)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/chats/19:bogus@thread.v2/members',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: '1', message: 'NotFound' } }), {
            status: 404,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-chat-members'];
    if (!cmd) throw new Error('list-chat-members not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { chatId: '19:bogus@thread.v2' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('NotFound: Microsoft Teams chat not found');
      expect(r.error.message).toContain('19:bogus@thread.v2');
      expect(r.error.message).toContain('list-chats');
    }
  });

  it('get-team-channel rewrites the opaque `1: NotFound` error to a clear channel-id hint (audit v1.0.0 B2)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/teams/tm1/channels/19:bogus@thread.skype',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: '1', message: 'NotFound' } }), {
            status: 404,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['get-team-channel'];
    if (!cmd) throw new Error('get-team-channel not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { teamId: 'tm1', channelId: '19:bogus@thread.skype' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('NotFound: Microsoft Teams channel not found');
      expect(r.error.message).toContain('19:bogus@thread.skype');
      expect(r.error.message).toContain('list-team-channels');
    }
  });

  it('list-sharepoint-site-onenote-notebooks rewrites Graph error 10008 (5k-item OneNote limit) to a one-line actionable summary — audit round-8 H2', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/sites/s1/onenote/notebooks',
        method: 'GET',
        response: () =>
          new Response(
            JSON.stringify({
              error: {
                code: '10008',
                message:
                  'The OneNote service is currently unavailable for this tenant because the SharePoint site collection has too many items. Browse to this page for more information: https://blogs.msdn.microsoft.com/onenotedev/2016/09/11/',
              },
            }),
            { status: 503, headers: { 'content-type': 'application/json' } }
          ),
      },
    ]);
    const cmd = cmdMap['list-sharepoint-site-onenote-notebooks'];
    if (!cmd) throw new Error('list-sharepoint-site-onenote-notebooks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { siteId: 's1' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('5000-item limit');
      expect(r.error.message).toContain('admin');
      expect(r.error.message).not.toContain('msdn.microsoft.com');
      expect(r.error.code).toBe('cli_rewrite_onenote_5k_limit');
    }
  });

  it('get-my-manager returns `{ manager: null, note }` (not bare null) when Graph 404s Request_ResourceNotFound — audit round-8 H1', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/manager',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'Request_ResourceNotFound', message: 'Resource not found.' } }), {
            status: 404,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['get-my-manager'];
    if (!cmd) throw new Error('get-my-manager not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, {});
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { manager: null; note: string };
      expect(v.manager).toBeNull();
      expect(typeof v.note).toBe('string');
      expect(v.note).toContain('no manager');
    }
  });

  it('list-incomplete-todo-tasks passes through RequestBroker--ParseUri unchanged when neither --select nor --orderby is set (no false rewrite)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: "https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks?$filter=status ne 'completed'",
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-incomplete-todo-tasks'];
    if (!cmd) throw new Error('list-incomplete-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('RequestBroker--ParseUri');
      expect(r.error.message).not.toContain('Drop `title`');
    }
  });

  it('list-incomplete-todo-tasks rewrites --orderby title with the parser-quirk hint (audit round-8 §1.1)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: "https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks?$filter=status ne 'completed'&$orderby=title%20asc",
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-incomplete-todo-tasks'];
    if (!cmd) throw new Error('list-incomplete-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', orderby: 'title asc' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('Graph rejected --orderby=title asc');
      expect(r.error.message).toContain('sorting on `title` is unsupported');
    }
  });

  it('list-incomplete-todo-tasks rewrites the title-quirk RequestBroker--ParseUri error like its sibling (audit round-8 §1.1)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: "https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks?$filter=status ne 'completed'&$select=id%2Ctitle",
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['list-incomplete-todo-tasks'];
    if (!cmd) throw new Error('list-incomplete-todo-tasks not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', select: 'id,title' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'api_error') {
      expect(r.error.message).toContain('Graph rejected --select=id,title');
      expect(r.error.message).toContain('Drop `title` from --select');
    }
  });

  it('list-incomplete-todo-tasks rejects --filter with a pointer at list-todo-tasks (audit round-6 §2.7)', async () => {
    const cmd = cmdMap['list-incomplete-todo-tasks'];
    if (!cmd) throw new Error('list-incomplete-todo-tasks not registered');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { todoTaskListId: 'tasks', filter: "subject eq 'x'" });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.message).toContain('--filter is not supported on list-incomplete-todo-tasks');
      expect(r.error.message).toContain('list-todo-tasks');
    }
  });

  it('list-incomplete-planner-tasks rejects --filter with a pointer at list-planner-tasks (audit round-6 §2.7)', async () => {
    const cmd = cmdMap['list-incomplete-planner-tasks'];
    if (!cmd) throw new Error('list-incomplete-planner-tasks not registered');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { filter: 'priority eq 1' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.message).toContain('--filter is not supported on list-incomplete-planner-tasks');
      expect(r.error.message).toContain('list-planner-tasks');
    }
  });

  it('download-onedrive-file-content returns plain-text source extensions inline as `{contentType: "text/plain", text}` (no 33% base64 bloat — audit v1.0.0 §bug-3)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'README.md', size: 5 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response(new TextEncoder().encode('# hi') as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-onedrive-file-content'];
    if (!cmd) throw new Error('download-onedrive-file-content not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string; base64?: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toBe('# hi');
      expect(v.base64).toBeUndefined();
    }
  });

  it('download-drive-item-version-content inlines the historical-version bytes via the M365ChatClient-elevated path', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/v3.bin' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/v3.bin',
        method: 'GET',
        response: () =>
          new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-content'];
    if (!cmd) throw new Error('download-drive-item-version-content not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/octet-stream');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-as-pdf converts an Office source via Graph ?format=pdf and inlines the PDF bytes (CDN redirect followed internally)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'q3.docx', size: 9 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/content?format=pdf',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/q3.pdf' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/q3.pdf',
        method: 'GET',
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-as-pdf short-circuits plain-text source extensions to `{contentType: "text/plain", size, text}` for envelope parity with download-onedrive-file-content (audit round-7 B5)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'README.md', size: 4 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response('hi', { status: 200, headers: { 'content-type': 'text/markdown' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; text: string; passthrough: true; note: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toBe('hi');
      expect(v.passthrough).toBe(true);
    }
  });

  it('download-drive-item-as-markdown converts a docx via the local mammoth pipeline (no Graph format=html call)', async () => {
    const docxBytes = await buildSampleDocx();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'q3.docx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/content',
        method: 'GET',
        response: () => new Response(docxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Sample Heading');
    }
  });

  it('download-drive-item-as-markdown short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'notes.txt' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('plain');
    }
  });

  it('download-drive-item-as-pdf propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'not found' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-version-as-pdf converts a non-current version through Graph ?format=pdf and inlines the PDF bytes (CDN redirect followed via the M365ChatClient-elevated path)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'budget.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content?format=pdf',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/v3.pdf' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/v3.pdf',
        method: 'GET',
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-version-as-markdown converts a non-current xlsx version via the local sheetjs pipeline', async () => {
    const xlsxBytes = buildSampleXlsx();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'budget.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content',
        method: 'GET',
        response: () => new Response(xlsxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('## Sheet1');
    }
  });

  it('download-drive-item-version-as-markdown short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'notes.md' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('# v2', { status: 200, headers: { 'content-type': 'text/markdown' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('# v2');
    }
  });

  it('download-drive-item-version-as-pdf short-circuits to a raw bytes download for plain-text source extensions (re-encodes text body as base64)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'log.log' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('line', { status: 200, headers: { 'content-type': 'text/plain' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('text/plain');
      expect(atob(v.base64)).toBe('line');
    }
  });

  it('download-drive-item-as-pdf short-circuits to raw download when the source itself is already a pdf (avoids the format=pdf 406 InputFormatNotSupported) and inlines the bytes', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf', method: 'GET', response: Response.json({ name: 'report.pdf' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf/content',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/report.pdf' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/report.pdf',
        method: 'GET',
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-version-as-pdf short-circuits to raw download for a pdf source (same reason as the non-versioned variant) and inlines the bytes', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf', method: 'GET', response: Response.json({ name: 'archive.pdf' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf/versions/2.0/content',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/v2.pdf' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/v2.pdf',
        method: 'GET',
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iPdf', versionId: '2.0' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-version-as-pdf propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-pdf'];
    if (!cmd) throw new Error('download-drive-item-version-as-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-version-as-markdown propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-version-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-as-markdown propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('extract-sharepoint-links-in-mail surfaces every SP URL in the body and resolves each via /shares/{token}/driveItem', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m1',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'Q3 deck',
            body: {
              content:
                '<p>See <a href="https://contoso.sharepoint.com/sites/Marketing/Q3.docx">deck</a> and <a href="https://contoso.sharepoint.com/sites/Marketing/notes.txt">notes</a>.</p>',
            },
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: () => Response.json({ id: 'i-q3', name: 'Q3.docx', webUrl: 'https://contoso.sharepoint.com/sites/Marketing/Q3.docx', parentReference: { driveId: 'd1' } }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: () => Response.json({ id: 'i-notes', name: 'notes.txt', webUrl: 'https://contoso.sharepoint.com/sites/Marketing/notes.txt', parentReference: { driveId: 'd1' } }),
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as {
        messageId: string;
        subject?: string;
        links: Array<{ url: string; driveId?: string; itemId?: string; name?: string }>;
        truncated: boolean;
        skippedCount: number;
      };
      expect(v.messageId).toBe('m1');
      expect(v.subject).toBe('Q3 deck');
      expect(v.truncated).toBe(false);
      expect(v.skippedCount).toBe(0);
      expect(v.links).toHaveLength(2);
      expect(v.links[0]?.driveId).toBe('d1');
      expect(v.links[0]?.itemId).toBe('i-q3');
      expect(v.links[0]?.name).toBe('Q3.docx');
      expect(v.links[1]?.itemId).toBe('i-notes');
    }
  });

  it('extract-sharepoint-links-in-mail caps the response at 25 links and reports `truncated` + `skippedCount` (Hardening #4)', async () => {
    const links = Array.from({ length: 30 }, (_, i) => `<a href="https://contoso.sharepoint.com/sites/X/file${i}.docx">f${i}</a>`).join('');
    const handlers = [
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBig',
        method: 'GET',
        response: (): Response => Response.json({ subject: 'Many links', body: { content: `<p>${links}</p>` } }),
      },
      ...Array.from({ length: 25 }, (_, i) => ({
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: (): Response => Response.json({ id: `i-${i}`, name: `file${i}.docx`, parentReference: { driveId: 'd1' } }),
      })),
    ];
    const fetchFn = stagedFetch(handlers);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mBig' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { links: Array<unknown>; truncated: boolean; skippedCount: number };
      expect(v.links).toHaveLength(25);
      expect(v.truncated).toBe(true);
      expect(v.skippedCount).toBe(5);
    }
  });

  it('extract-sharepoint-links-in-mail captures per-link resolve errors instead of failing the whole call', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m2',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'mixed',
            body: { content: '<a href="https://contoso.sharepoint.com/good.docx">a</a> <a href="https://contoso.sharepoint.com/bad.docx">b</a>' },
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: () => Response.json({ id: 'i-good', name: 'good.docx', parentReference: { driveId: 'd1' } }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'access denied' } }), { status: 403, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm2' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { links: Array<{ url: string; itemId?: string; error?: string }> };
      expect(v.links).toHaveLength(2);
      expect(v.links[0]?.itemId).toBe('i-good');
      expect(v.links[1]?.error).toContain('access denied');
    }
  });

  it('extract-sharepoint-links-in-mail returns an empty links array when the body has no SharePoint URLs', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m3',
        method: 'GET',
        response: () => Response.json({ subject: 'plain', body: { content: '<p>just text</p>' } }),
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm3' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { links: Array<unknown>; truncated: boolean };
      expect(v.links).toHaveLength(0);
      expect(v.truncated).toBe(false);
    }
  });

  it('extract-sharepoint-links-in-mail propagates an err from the message GET unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNope',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'not found' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mNope' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('extract-sharepoint-links-in-mail handles a message with no body field', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mEmpty',
        method: 'GET',
        response: () => Response.json({ subject: 'no body' }),
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mEmpty' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { links: Array<unknown> };
      expect(v.links).toHaveLength(0);
    }
  });

  it('extract-sharepoint-links-in-mail surfaces a non-api_error err type with a labelled message', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNetErr',
        method: 'GET',
        response: () => Response.json({ subject: 'x', body: { content: '<a href="https://contoso.sharepoint.com/a.docx">a</a>' } }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/shares/u!',
        method: 'GET',
        response: () => {
          throw new Error('boom');
        },
      },
    ]);
    const cmd = cmdMap['extract-sharepoint-links-in-mail'];
    if (!cmd) throw new Error('extract-sharepoint-links-in-mail not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mNetErr' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { links: Array<{ error?: string }> };
      expect(v.links[0]?.error).toContain('network_error');
      expect(v.links[0]?.error).toContain('boom');
    }
  });

  it('convert-mail-to-markdown renders a single email with headers + body via turndown', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m1',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'Q3',
            from: { emailAddress: { name: 'Alice', address: 'alice@contoso.com' } },
            toRecipients: [{ emailAddress: { address: 'bob@contoso.com' } }],
            receivedDateTime: '2026-04-30T08:00:00Z',
            body: { contentType: 'html', content: '<p>Hi <strong>team</strong>.</p>' },
            attachments: [],
          }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('**Subject:** Q3');
      expect(v.text).toContain('**From:** Alice <alice@contoso.com>');
      expect(v.text).toContain('Hi **team**.');
    }
  });

  it('convert-mail-to-markdown embeds inline image attachments as data: URIs (Hardening #1)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m2',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'with logo',
            body: { contentType: 'html', content: '<p>Logo: <img src="cid:logo123" alt="logo"></p>' },
            attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', isInline: true, contentId: 'logo123', contentType: 'image/png', contentBytes: 'iVBORw0=' }],
          }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm2' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('data:image/png;base64,iVBORw0=');
      expect(v.text).not.toContain('cid:logo123');
    }
  });

  it('convert-mail-to-markdown does NOT embed non-image inline attachments (Hardening #1 verified)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m3',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'sneaky',
            body: { contentType: 'html', content: '<p>X: <img src="cid:evil"></p>' },
            attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', isInline: true, contentId: 'evil', contentType: 'text/html', contentBytes: 'PHNjcmlwdD4=' }],
          }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm3' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).not.toContain('PHNjcmlwdD4=');
      expect(v.text).not.toContain('data:text/html');
    }
  });

  it('convert-mail-to-markdown handles plain-text bodies (no turndown round-trip)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m4',
        method: 'GET',
        response: () => Response.json({ subject: 'plain', body: { contentType: 'text', content: 'Hello\nworld' }, attachments: [] }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm4' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('Hello');
    }
  });

  it('convert-mail-attachment-to-pdf uploads a fileAttachment, converts via ?format=pdf, follows the CDN redirect internally to inline the PDF bytes, then deletes the temp item', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/a1')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx-bytes') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i1', name: 'plan-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/plan.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/plan.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i1') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
    expect(calls.some((c) => c.method === 'DELETE')).toBe(true);
    expect(calls.some((c) => c.url.endsWith('/content?format=pdf'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf short-circuits to a raw-bytes envelope for plain-text source extensions', async () => {
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith('/attachments/aText')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README.md', contentBytes: btoa('# Hello') });
      }
      throw new Error(`unexpected fetch ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; size: number; base64: string; note: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.note).toContain('pre-checked source');
      expect(atob(v.base64)).toBe('# Hello');
    }
  });

  it('convert-mail-attachment-to-pdf short-circuits a pdf fileAttachment to its raw bytes (no upload-convert dance, no Graph format=pdf call)', async () => {
    const calls: Array<{ method: string; url: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ method: init?.method ?? 'GET', url });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: btoa('%PDF-fake') });
      }
      throw new Error(`unexpected fetch ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string; note: string };
      expect(v.contentType).toBe('application/pdf');
      expect(v.note).toContain('pre-checked source');
      expect(atob(v.base64)).toBe('%PDF-fake');
    }
    expect(calls.some((c) => c.method === 'PUT')).toBe(false);
    expect(calls.some((c) => c.url.includes('format=pdf'))).toBe(false);
  });

  it('convert-mail-attachment-to-pdf short-circuits a pdf referenceAttachment to its raw bytes via /content (no format=pdf), follows the CDN redirect internally', async () => {
    let formatPdfCalled = false;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/report.pdf' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-pdf', name: 'report.pdf', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-pdf/content')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/report.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/report.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.includes('format=pdf')) {
        formatPdfCalled = true;
        return new Response(null, { status: 500 });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
    expect(formatPdfCalled).toBe(false);
  });

  it('convert-mail-attachment-to-pdf resolves a referenceAttachment via /shares/{token}/driveItem, converts in place, and inlines the PDF bytes (CDN redirect followed internally)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRef')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/q3.docx' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-q3', name: 'q3.docx', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-q3/content?format=pdf')) {
        return Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/q3.pdf' });
      }
      if (url === 'https://contoso.sharepoint.com/cdn/q3.pdf') {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]) as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRef' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('convert-mail-attachment-to-pdf rejects an image fileAttachment with a friendly hint pointing at get-mail-attachment + a vision-capable model (audit v1.0.0 §2.4)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPng')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'screenshot.png', contentBytes: btoa('fake-png-bytes') });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPng' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('png attachment is an image');
      expect(result.error.message).toContain('get-mail-attachment');
      expect(result.error.message).toContain('vision-capable model');
    }
  });

  it('convert-mail-attachment-to-pdf rejects an image referenceAttachment with the same friendly hint (no upload-then-fail dance)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefImg')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/diagram.svg' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-svg', name: 'diagram.svg', parentReference: { driveId: 'd1' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefImg' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('svg attachment is an image');
    }
  });

  it('convert-mail-attachment-to-pdf rejects itemAttachment with a clear unsupported error pointing at the markdown variant', async () => {
    const fetchFn: FetchFn = async () =>
      Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: { '@odata.type': '#microsoft.graph.message', subject: 'embedded' } });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aItem' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(400);
      expect(result.error.message).toContain('convert-mail-attachment-to-markdown');
    }
  });

  it('convert-mail-attachment-to-pdf returns api_error when @odata.type is missing (FIX #3 discriminator guard)', async () => {
    const fetchFn: FetchFn = async () => Response.json({ name: 'no-type.docx' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('missing @odata.type discriminator');
    }
  });

  it('convert-mail-attachment-to-pdf returns api_error when the @odata.type is unknown', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.weirdNewType' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aWeird' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unsupported attachment type');
    }
  });

  it('convert-mail-attachment-to-markdown renders an itemAttachment (message) without any Graph conversion call', async () => {
    let conversionCalled = false;
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aMsg')) {
        return Response.json({
          '@odata.type': '#microsoft.graph.itemAttachment',
          item: {
            '@odata.type': '#microsoft.graph.message',
            subject: 'Re: Q3',
            from: { emailAddress: { address: 'alice@x' } },
            body: { contentType: 'html', content: '<p>looks good</p>' },
          },
        });
      }
      if (url.includes('?format=html')) conversionCalled = true;
      return Response.json({});
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aMsg' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('**Subject:** Re: Q3');
      expect(v.text).toContain('looks good');
    }
    expect(conversionCalled).toBe(false);
  });

  it('convert-mail-attachment-to-markdown renders an itemAttachment (event)', async () => {
    const fetchFn: FetchFn = async () =>
      Response.json({
        '@odata.type': '#microsoft.graph.itemAttachment',
        item: {
          '@odata.type': '#microsoft.graph.event',
          subject: 'Quarterly Review',
          start: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' },
          end: { dateTime: '2026-05-01T10:00:00', timeZone: 'UTC' },
        },
      });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aEvent' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('**Subject:** Quarterly Review');
      expect(v.text).toContain('**Start:** 2026-05-01T09:00:00 UTC');
    }
  });

  it('convert-mail-attachment-to-markdown rejects an itemAttachment whose inner @odata.type is unknown', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: { '@odata.type': '#microsoft.graph.weird' } });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aWeird' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('unsupported embedded item type');
    }
  });

  it('convert-mail-attachment-to-markdown short-circuits a plain-text fileAttachment to a raw envelope', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README.md', contentBytes: btoa('# Hi') });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; note?: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.note).toContain('plain-text source');
    }
  });

  it('convert-mail-attachment-to-markdown resolves a referenceAttachment via /shares/{token}/driveItem and runs the local docx pipeline', async () => {
    const docxBytes = await buildSampleDocx();
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRef')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/q3.docx' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-q3', name: 'q3.docx', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-q3/content')) {
        return new Response(docxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRef' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('# Sample Heading');
    }
  });

  it('convert-mail-attachment-to-markdown short-circuits referenceAttachment to a raw download for plain-text source', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefText')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/notes.txt' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-text', name: 'notes.txt', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-text/content')) {
        return new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text?: string };
      expect(v.text).toBe('plain');
    }
  });

  it('convert-mail-attachment-to-markdown rejects a referenceAttachment with no sourceUrl', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing sourceUrl');
  });

  it('convert-mail-attachment-to-markdown rejects a referenceAttachment whose /shares resolution lacks driveId', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefBad2')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x.docx' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'no-drive' });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad2' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing id or driveId');
  });

  it('convert-mail-attachment-to-markdown renders an itemAttachment (contact)', async () => {
    const fetchFn: FetchFn = async () =>
      Response.json({
        '@odata.type': '#microsoft.graph.itemAttachment',
        item: { '@odata.type': '#microsoft.graph.contact', displayName: 'Alice', emailAddresses: [{ address: 'alice@x' }] },
      });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aContact' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('**Name:** Alice');
      expect(v.text).toContain('**Emails:** alice@x');
    }
  });

  it('convert-mail-attachment-to-markdown rejects an itemAttachment with no inner item field', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment' });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aBare' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing inner item');
  });

  it('convert-mail-attachment-to-markdown rejects an itemAttachment whose inner item lacks @odata.type', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: { subject: 'no type' } });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aNoType' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('itemAttachment.item missing @odata.type');
  });

  it('convert-mail-attachment-to-markdown returns api_error when @odata.type is missing on the outer attachment (FIX #3)', async () => {
    const fetchFn: FetchFn = async () => Response.json({ name: 'no-type.docx' });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing @odata.type discriminator');
  });

  it('convert-mail-attachment-to-markdown returns api_error when the @odata.type is unknown', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.weirdNewType' });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aWeird' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('unsupported attachment type');
  });

  it('convert-mail-attachment-to-markdown rejects an image attachment with a hint pointing at get-mail-attachment, not the PDF converter', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'logo.png', contentBytes: 'AAAA' });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aImg' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('image');
      expect(result.error.message).toContain('get-mail-attachment');
    }
  });

  it('convert-mail-attachment-to-pdf rejects a referenceAttachment with no sourceUrl', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' });
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing sourceUrl');
  });

  it('convert-mail-attachment-to-pdf rejects a referenceAttachment whose resolved driveItem lacks ids', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefBad2')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x.docx' });
      }
      if (url.includes('/shares/u!')) return Response.json({ id: 'no-drive' });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefBad2' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('missing id or driveId');
  });

  it('convert-mail-attachment-to-pdf short-circuits a referenceAttachment to raw bytes for plain-text source', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefText')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/notes.txt' });
      }
      if (url.includes('/shares/u!')) return Response.json({ id: 'i-text', name: 'notes.txt', parentReference: { driveId: 'd1' } });
      if (url.endsWith('/drives/d1/items/i-text/content')) return new Response('plain', { status: 200, headers: { 'content-type': 'text/plain' } });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aRefText' });
    expect(result.ok).toBe(true);
  });

  it('convert-mail-attachment-to-pdf returns api_error when upload returns no driveItem id', async () => {
    // Distinct attachment id so the fake's URL-suffix check differs from the markdown variant's test above.
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith('/attachments/aFilePdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx-bytes') });
      }
      if (init?.method === 'PUT') return Response.json({ name: 'no-id' });
      throw new Error('unexpected');
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aFilePdf' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('upload returned no driveItem id');
  });

  it('convert-mail-attachment-to-markdown decodes a docx fileAttachment locally and runs it through mammoth — no upload, no Graph format=html', async () => {
    const docxBytes = await buildSampleDocx();
    let putCalled = false;
    let deleteCalled = false;
    const fetchFn: FetchFn = async (url, init) => {
      if (url.endsWith('/attachments/aFile')) {
        const b64 = btoa(String.fromCharCode(...docxBytes));
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: b64 });
      }
      if (init?.method === 'PUT') {
        putCalled = true;
        return new Response(null, { status: 200 });
      }
      if (init?.method === 'DELETE') {
        deleteCalled = true;
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected fetch ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aFile' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Sample Heading');
    }
    expect(putCalled).toBe(false);
    expect(deleteCalled).toBe(false);
  });

  it('convert-mail-attachment-to-markdown errs with the pptx-specific PDF hint for a pptx fileAttachment', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPptx')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentBytes: btoa('zzz') });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPptx' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('pptx attachment');
      expect(result.error.message).toContain('convert-mail-attachment-to-pdf');
    }
  });

  it('convert-mail-attachment-to-markdown errs on a PDF fileAttachment with a no-PDF→markdown-path hint pointing at vision models or external tools (audit v1.0.0 §bug-6 — the convert-…-to-pdf fallback is circular for PDF inputs)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: btoa('zzz') });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPdf' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('pdf attachment cannot be converted to markdown');
      expect(result.error.message).toContain('vision-capable model');
      expect(result.error.message).toContain('external PDF→text tool');
    }
  });

  it('convert-mail-attachment-to-markdown errs with `<no-extension>` placeholder for fileAttachments without a recognizable extension', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aNoExt')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README', contentBytes: btoa('zzz') });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aNoExt' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('<no-extension>');
    }
  });

  it('get-onenote-page-as-markdown turns Graph-returned HTML into a markdown envelope', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/onenote/pages/p1/content',
        method: 'GET',
        response: () => new Response('<h1>Meeting</h1>', { status: 200, headers: { 'content-type': 'text/html' } }),
      },
    ]);
    const cmd = cmdMap['get-onenote-page-as-markdown'];
    if (!cmd) throw new Error('get-onenote-page-as-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { onenotePageId: 'p1' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Meeting');
    }
  });
});

type CommandFixture = { readonly name: string; readonly params: Record<string, string>; readonly responseBody?: unknown };

const allCommandFixtures: CommandFixture[] = [
  { name: 'list-drives', params: {} },
  { name: 'get-drive-root-item', params: { driveId: 'd1' } },
  { name: 'list-folder-files', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'download-onedrive-file-content', params: { driveId: 'd1', itemId: 'i1' }, responseBody: { contentType: 'application/octet-stream', size: 5, base64: 'JVBERi0=' } },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' } },
  {
    name: 'download-drive-item-version-content',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    responseBody: { contentType: 'application/octet-stream', size: 5, base64: 'JVBERi0=' },
  },
  { name: 'download-drive-item-as-pdf', params: { driveId: 'd1', itemId: 'i1' }, responseBody: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } },
  {
    name: 'download-drive-item-version-as-pdf',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    responseBody: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' },
  },
  { name: 'search-onedrive-files', params: { driveId: 'd1', query: 'report' } },
  { name: 'search-my-documents', params: { query: 'budget' } },
  { name: 'get-excel-range', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1' } },
  { name: 'list-excel-worksheets', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-excel-tables', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-excel-table', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' } },
  { name: 'list-excel-table-rows', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' } },
  { name: 'get-drive-delta', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'search-sharepoint-sites-by-name', params: { query: 'marketing' } },
  { name: 'get-sharepoint-site', params: { siteId: 's1' } },
  { name: 'list-sharepoint-site-drives', params: { siteId: 's1' } },
  { name: 'get-sharepoint-site-drive-by-id', params: { siteId: 's1', driveId: 'd1' } },
  { name: 'list-sharepoint-site-lists', params: { siteId: 's1' } },
  { name: 'get-sharepoint-site-list', params: { siteId: 's1', listId: 'l1' } },
  { name: 'list-sharepoint-site-list-items', params: { siteId: 's1', listId: 'l1' } },
  { name: 'get-sharepoint-site-list-item', params: { siteId: 's1', listId: 'l1', listItemId: 'li1' } },
  { name: 'get-sharepoint-site-by-path', params: { hostname: 'contoso.sharepoint.com', path: '/sites/Marketing' } },
  { name: 'list-todo-task-lists', params: {} },
  { name: 'list-todo-tasks', params: { todoTaskListId: 'tl1' } },
  { name: 'list-incomplete-todo-tasks', params: { todoTaskListId: 'tl1' } },
  { name: 'get-todo-task', params: { todoTaskListId: 'tl1', todoTaskId: 't1' } },
  { name: 'list-todo-linked-resources', params: { todoTaskListId: 'tl1', todoTaskId: 't1' } },
  { name: 'list-planner-plans', params: {} },
  { name: 'list-planner-tasks', params: {} },
  { name: 'list-incomplete-planner-tasks', params: {} },
  { name: 'get-planner-plan', params: { plannerPlanId: 'p1' } },
  { name: 'list-plan-tasks', params: { plannerPlanId: 'p1' } },
  { name: 'get-planner-task', params: { plannerTaskId: 't1' } },
  { name: 'get-planner-task-details', params: { plannerTaskId: 't1' } },
  { name: 'list-plan-buckets', params: { plannerPlanId: 'p1' } },
  { name: 'get-planner-bucket', params: { plannerBucketId: 'b1' } },
  { name: 'list-mail-messages', params: {} },
  { name: 'list-mail-folders', params: {} },
  { name: 'list-mail-child-folders', params: { mailFolderId: 'f1' } },
  { name: 'list-mail-folder-messages', params: { mailFolderId: 'f1' } },
  { name: 'get-mail-message', params: { messageId: 'm1' } },
  { name: 'list-mail-attachments', params: { messageId: 'm1' } },
  { name: 'get-mail-attachment', params: { messageId: 'm1', attachmentId: 'a1' } },
  { name: 'list-mail-rules', params: { mailFolderId: 'f1' } },
  { name: 'get-mailbox-settings', params: {} },
  { name: 'search-mail-messages', params: { query: 'invoice' } },
  { name: 'extract-sharepoint-links-in-mail', params: { messageId: 'm1' } },
  { name: 'convert-mail-to-markdown', params: { messageId: 'm1' } },
  { name: 'list-onenote-notebooks', params: {} },
  { name: 'list-onenote-notebook-sections', params: { notebookId: 'n1' } },
  { name: 'list-all-onenote-sections', params: {} },
  { name: 'list-onenote-section-pages', params: { onenoteSectionId: 's1' } },
  { name: 'get-onenote-page-content', params: { onenotePageId: 'p1' } },
  { name: 'search-onenote-pages', params: { titleSubstring: 'meeting' } },
  { name: 'get-current-user', params: {} },
  { name: 'get-my-profile-photo', params: {}, responseBody: { contentType: 'image/jpeg', size: 5, base64: 'JVBERi0=' } },
  { name: 'list-calendar-events', params: {} },
  { name: 'get-calendar-event', params: { eventId: 'e1' } },
  { name: 'list-specific-calendar-events', params: { calendarId: 'c1' } },
  { name: 'get-specific-calendar-event', params: { calendarId: 'c1', eventId: 'e1' } },
  { name: 'list-calendar-view', params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-specific-calendar-view', params: { calendarId: 'c1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-calendar-event-instances', params: { calendarId: 'c1', eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-calendars', params: {} },
  { name: 'list-calendar-events-delta', params: {} },
  { name: 'list-calendar-view-delta', params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-chat-members', params: { chatId: 'ch1' } },
  { name: 'list-joined-teams', params: {} },
  { name: 'get-team', params: { teamId: 'tm1' } },
  { name: 'list-team-channels', params: { teamId: 'tm1' } },
  { name: 'get-team-channel', params: { teamId: 'tm1', channelId: 'ch1' } },
  { name: 'list-chats', params: {} },
  { name: 'get-chat', params: { chatId: 'ch1' } },
  { name: 'list-my-direct-reports', params: {} },
  { name: 'list-user-direct-reports', params: { userId: 'alice@contoso.com' } },
  { name: 'list-my-memberships', params: {} },
  { name: 'get-my-manager', params: {} },
  { name: 'get-user-manager', params: { userId: 'alice@contoso.com' } },
  { name: 'list-relevant-people', params: {} },
  { name: 'list-groups', params: {} },
  { name: 'get-group', params: { groupId: 'g1' } },
  { name: 'list-group-members', params: { groupId: 'g1' } },
  { name: 'list-group-owners', params: { groupId: 'g1' } },
  { name: 'list-group-events', params: { groupId: 'g1' } },
  { name: 'list-group-calendar-view', params: { groupId: 'g1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-group-conversations', params: { groupId: 'g1' } },
  { name: 'list-group-threads', params: { groupId: 'g1' } },
  { name: 'get-mail-message-mime', params: { messageId: 'm1' }, responseBody: { contentType: 'message/rfc822', size: 5, base64: 'JVBERi0=' } },
  { name: 'list-mail-folder-messages-delta', params: { mailFolderId: 'inbox' } },
  { name: 'list-shared-mailbox-messages', params: { userId: 'shared@contoso.com' } },
  { name: 'list-shared-mailbox-folder-messages', params: { userId: 'shared@contoso.com', mailFolderId: 'inbox' } },
  { name: 'get-shared-mailbox-message', params: { userId: 'shared@contoso.com', messageId: 'm1' } },
  { name: 'list-conversation-messages', params: { conversationId: 'AAQkAD-conv-1' } },
  { name: 'list-focused-inbox-overrides', params: {} },
  { name: 'list-outlook-categories', params: {} },
  { name: 'list-shared-calendar-events', params: { userId: 'colleague@contoso.com' } },
  { name: 'list-shared-calendar-view', params: { userId: 'colleague@contoso.com', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' } },
  { name: 'list-sharepoint-list-columns', params: { siteId: 's1', listId: 'l1' } },
  { name: 'get-sharepoint-list-column', params: { siteId: 's1', listId: 'l1', columnId: 'Title' } },
  { name: 'list-sharepoint-site-onenote-notebooks', params: { siteId: 's1' } },
  { name: 'list-sharepoint-site-onenote-notebook-sections', params: { siteId: 's1', notebookId: 'nb1' } },
  { name: 'list-sharepoint-site-onenote-section-pages', params: { siteId: 's1', onenoteSectionId: 'sec1' } },
  { name: 'get-sharepoint-site-onenote-page-content', params: { siteId: 's1', onenotePageId: 'p1' } },
  { name: 'list-drive-item-thumbnails', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-excel-used-range', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' } },
  { name: 'list-rooms', params: {} },
  { name: 'list-room-lists', params: {} },
  { name: 'list-trending-insights', params: {} },
  { name: 'list-recent-files', params: {} },
  { name: 'list-shared-with-me', params: {} },
  { name: 'list-recently-used-insights', params: {} },
  { name: 'list-shared-insights', params: {} },
  { name: 'get-organization', params: {} },
  { name: 'list-mail-folders-delta', params: {} },
  { name: 'get-channel-files-folder', params: { teamId: 'tm1', channelId: 'ch1' } },
  { name: 'get-drive-item-list-item', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-drive-item-analytics', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-team-installed-apps', params: { teamId: 'tm1' } },
  { name: 'list-calendar-groups', params: {} },
  { name: 'list-calendar-group-calendars', params: { calendarGroupId: 'cg1' } },
  { name: 'get-my-calendar', params: {} },
  { name: 'list-site-columns', params: { siteId: 's1' } },
  { name: 'list-site-content-types', params: { siteId: 's1' } },
  { name: 'list-sharepoint-site-pages', params: { siteId: 's1' } },
  { name: 'list-excel-defined-names', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-excel-worksheet-charts', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' } },
  { name: 'microsoft-search-query', params: { query: 'q3 budget' } },
  { name: 'get-drive-special-folder', params: { folderName: 'documents' } },
  { name: 'get-drive-root-delta', params: {} },
  { name: 'list-followed-drive-items', params: {} },
  { name: 'get-drive-item-created-by-user', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-drive-item-last-modified-by-user', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'get-site-analytics', params: { siteId: 's1' } },
  { name: 'list-sharepoint-list-item-versions', params: { siteId: 's1', listId: 'l1', listItemId: '12' } },
  { name: 'get-mail-rule', params: { mailFolderId: 'inbox', messageRuleId: 'r1' } },
  { name: 'list-excel-comments', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-excel-worksheet-pivot-tables', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' } },
  { name: 'list-sensitivity-labels', params: {} },
  { name: 'list-my-transitive-memberships', params: {} },
  { name: 'get-team-primary-channel', params: { teamId: 'tm1' } },
  { name: 'list-todo-tasks-delta', params: { todoTaskListId: 'tl1' } },
  { name: 'next-page', params: { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' } },
];

describe('all commands schema acceptance', () => {
  it.each(allCommandFixtures)('$name accepts valid params', async ({ name, params, responseBody }) => {
    const result = await callCommand(name, params, responseBody ?? { ok: true });
    expect(result.ok).toBe(true);
  });
});

describe('command schema rejection', () => {
  const rejectCases: Array<{ name: string; params: Record<string, string> }> = [
    { name: 'get-drive-root-item', params: {} },
    { name: 'search-onedrive-files', params: { driveId: 'd1' } },
    { name: 'get-sharepoint-site', params: {} },
    { name: 'get-sharepoint-site-by-path', params: {} },
    { name: 'get-sharepoint-site-by-path', params: { hostname: 'contoso.sharepoint.com', path: 'sites/Marketing' } },
    { name: 'get-todo-task', params: { todoTaskListId: 'tl1' } },
    { name: 'get-planner-plan', params: {} },
    { name: 'get-mail-message', params: {} },
    { name: 'list-mail-child-folders', params: {} },
    { name: 'get-onenote-page-content', params: {} },
    { name: 'get-calendar-event', params: {} },
    { name: 'get-specific-calendar-event', params: { calendarId: 'c1' } },
    { name: 'list-team-channels', params: {} },
    { name: 'get-team-channel', params: { teamId: 'tm1' } },
    { name: 'download-onedrive-file-content', params: { driveId: 'd1' } },
    { name: 'download-drive-item-version-content', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'download-drive-item-as-pdf', params: { driveId: 'd1' } },
    { name: 'download-drive-item-as-markdown', params: { driveId: 'd1' } },
    { name: 'download-drive-item-version-as-pdf', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'download-drive-item-version-as-markdown', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'get-onenote-page-as-markdown', params: {} },
    { name: 'search-mail-messages', params: {} },
    { name: 'extract-sharepoint-links-in-mail', params: {} },
    { name: 'convert-mail-to-markdown', params: {} },
    { name: 'convert-mail-attachment-to-pdf', params: { messageId: 'm1' } },
    { name: 'convert-mail-attachment-to-markdown', params: { messageId: 'm1' } },
    { name: 'search-my-documents', params: {} },
    { name: 'list-calendar-view', params: {} },
    { name: 'list-calendar-view-delta', params: {} },
    { name: 'list-specific-calendar-view', params: { calendarId: 'c1' } },
    { name: 'list-calendar-event-instances', params: { calendarId: 'c1', eventId: 'e1' } },
    { name: 'search-onenote-pages', params: {} },
    { name: 'search-sharepoint-sites-by-name', params: {} },
    { name: 'list-incomplete-todo-tasks', params: {} },
    { name: 'next-page', params: {} },
    { name: 'next-page', params: { url: 'https://example.com/foo' } },
    { name: 'get-chat', params: {} },
    { name: 'list-user-direct-reports', params: {} },
    { name: 'get-user-manager', params: {} },
    { name: 'get-group', params: {} },
    { name: 'list-group-members', params: {} },
    { name: 'list-group-owners', params: {} },
    { name: 'list-group-events', params: {} },
    { name: 'list-group-calendar-view', params: {} },
    { name: 'list-group-conversations', params: {} },
    { name: 'list-group-threads', params: {} },
    { name: 'get-mail-message-mime', params: {} },
    { name: 'list-mail-folder-messages-delta', params: {} },
    { name: 'list-shared-mailbox-messages', params: {} },
    { name: 'list-shared-mailbox-folder-messages', params: {} },
    { name: 'get-shared-mailbox-message', params: {} },
    { name: 'list-conversation-messages', params: {} },
    { name: 'list-shared-calendar-events', params: {} },
    { name: 'list-shared-calendar-view', params: {} },
    { name: 'list-sharepoint-list-columns', params: {} },
    { name: 'get-sharepoint-list-column', params: {} },
    { name: 'list-sharepoint-site-onenote-notebooks', params: {} },
    { name: 'list-sharepoint-site-onenote-notebook-sections', params: {} },
    { name: 'list-sharepoint-site-onenote-section-pages', params: {} },
    { name: 'get-sharepoint-site-onenote-page-content', params: {} },
    { name: 'list-drive-item-thumbnails', params: {} },
    { name: 'get-excel-used-range', params: {} },
  ];

  it.each(rejectCases)('$name rejects missing required params as a validation_error Result', async ({ name, params }) => {
    const cmd = cmdMap[name];
    const fetchFn = fakeFetch({ ok: true });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, params);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});

const pathFixtures: Array<{ name: string; params: Record<string, string>; expectedPath: string }> = [
  { name: 'list-drives', params: {}, expectedPath: '/me/drives' },
  { name: 'get-drive-root-item', params: { driveId: 'd1' }, expectedPath: '/drives/d1/root' },
  { name: 'list-folder-files', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/children' },
  { name: 'download-onedrive-file-content', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/content' },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1' },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/permissions' },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/versions' },
  {
    name: 'download-drive-item-version-content',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    expectedPath: '/drives/d1/items/i1/versions/3.0/content',
  },
  {
    name: 'download-drive-item-as-pdf',
    params: { driveId: 'd1', itemId: 'i1' },
    expectedPath: '/drives/d1/items/i1/content?format=pdf',
  },
  {
    name: 'download-drive-item-as-markdown',
    params: { driveId: 'd1', itemId: 'i1' },
    // No-extension metadata response from fakeFetch makes the dispatcher fall into the unsupported branch
    // before any /content fetch — so the LAST URL hit is the metadata GET.
    expectedPath: '/drives/d1/items/i1',
  },
  {
    name: 'download-drive-item-version-as-pdf',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    expectedPath: '/drives/d1/items/i1/versions/3.0/content?format=pdf',
  },
  {
    name: 'download-drive-item-version-as-markdown',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    // Same as above — no-extension metadata short-circuits in the dispatcher.
    expectedPath: '/drives/d1/items/i1',
  },
  { name: 'search-onedrive-files', params: { driveId: 'd1', query: 'report' }, expectedPath: "/drives/d1/search(q='report')" },
  { name: 'search-my-documents', params: { query: 'budget' }, expectedPath: "/me/drive/search(q='budget')" },
  {
    name: 'get-excel-range',
    params: { driveId: 'd1', itemId: 'i1', worksheetId: 'ws1', address: 'A1' },
    expectedPath: "/drives/d1/items/i1/workbook/worksheets/ws1/range(address='A1')",
  },
  { name: 'list-excel-worksheets', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/worksheets' },
  { name: 'list-excel-tables', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/tables' },
  { name: 'get-excel-table', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' }, expectedPath: '/drives/d1/items/i1/workbook/tables/t1' },
  { name: 'list-excel-table-rows', params: { driveId: 'd1', itemId: 'i1', tableId: 't1' }, expectedPath: '/drives/d1/items/i1/workbook/tables/t1/rows' },
  { name: 'get-drive-delta', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/delta()' },
  { name: 'search-sharepoint-sites-by-name', params: { query: 'marketing' }, expectedPath: '/sites?search=marketing' },
  { name: 'get-sharepoint-site', params: { siteId: 's1' }, expectedPath: '/sites/s1' },
  { name: 'list-sharepoint-site-drives', params: { siteId: 's1' }, expectedPath: '/sites/s1/drives' },
  { name: 'get-sharepoint-site-drive-by-id', params: { siteId: 's1', driveId: 'd1' }, expectedPath: '/sites/s1/drives/d1' },
  { name: 'list-sharepoint-site-lists', params: { siteId: 's1' }, expectedPath: '/sites/s1/lists' },
  { name: 'get-sharepoint-site-list', params: { siteId: 's1', listId: 'l1' }, expectedPath: '/sites/s1/lists/l1' },
  { name: 'list-sharepoint-site-list-items', params: { siteId: 's1', listId: 'l1' }, expectedPath: '/sites/s1/lists/l1/items' },
  { name: 'get-sharepoint-site-list-item', params: { siteId: 's1', listId: 'l1', listItemId: 'li1' }, expectedPath: '/sites/s1/lists/l1/items/li1' },
  {
    name: 'get-sharepoint-site-by-path',
    params: { hostname: 'contoso.sharepoint.com', path: '/sites/Marketing' },
    expectedPath: '/sites/contoso.sharepoint.com:/sites/Marketing',
  },
  { name: 'list-todo-task-lists', params: {}, expectedPath: '/me/todo/lists' },
  { name: 'list-todo-tasks', params: { todoTaskListId: 'tl1' }, expectedPath: '/me/todo/lists/tl1/tasks' },
  { name: 'list-incomplete-todo-tasks', params: { todoTaskListId: 'tl1' }, expectedPath: "/me/todo/lists/tl1/tasks?$filter=status ne 'completed'" },
  { name: 'get-todo-task', params: { todoTaskListId: 'tl1', todoTaskId: 't1' }, expectedPath: '/me/todo/lists/tl1/tasks/t1' },
  { name: 'list-todo-linked-resources', params: { todoTaskListId: 'tl1', todoTaskId: 't1' }, expectedPath: '/me/todo/lists/tl1/tasks/t1/linkedResources' },
  { name: 'list-planner-plans', params: {}, expectedPath: '/me/planner/plans' },
  { name: 'list-planner-tasks', params: {}, expectedPath: '/me/planner/tasks' },
  { name: 'list-incomplete-planner-tasks', params: {}, expectedPath: '/me/planner/tasks?$filter=percentComplete ne 100' },
  { name: 'get-planner-plan', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1' },
  { name: 'list-plan-tasks', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1/tasks' },
  { name: 'get-planner-task', params: { plannerTaskId: 't1' }, expectedPath: '/planner/tasks/t1' },
  { name: 'get-planner-task-details', params: { plannerTaskId: 't1' }, expectedPath: '/planner/tasks/t1/details' },
  { name: 'list-plan-buckets', params: { plannerPlanId: 'p1' }, expectedPath: '/planner/plans/p1/buckets' },
  { name: 'get-planner-bucket', params: { plannerBucketId: 'b1' }, expectedPath: '/planner/buckets/b1' },
  { name: 'list-mail-messages', params: {}, expectedPath: '/me/messages' },
  { name: 'list-mail-folders', params: {}, expectedPath: '/me/mailFolders' },
  { name: 'list-mail-child-folders', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/childFolders' },
  { name: 'list-mail-folder-messages', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messages' },
  { name: 'get-mail-message', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1' },
  {
    name: 'list-mail-attachments',
    params: { messageId: 'm1' },
    expectedPath: '/me/messages/m1/attachments?$select=id%2Cname%2CcontentType%2Csize%2CisInline',
  },
  { name: 'get-mail-attachment', params: { messageId: 'm1', attachmentId: 'a1' }, expectedPath: '/me/messages/m1/attachments/a1' },
  { name: 'list-mail-rules', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messageRules' },
  { name: 'get-mailbox-settings', params: {}, expectedPath: '/me/mailboxSettings' },
  { name: 'search-mail-messages', params: { query: 'invoice' }, expectedPath: '/me/messages?$search="invoice"' },
  { name: 'extract-sharepoint-links-in-mail', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1?$select=subject,body' },
  { name: 'convert-mail-to-markdown', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1?$expand=attachments' },
  { name: 'list-onenote-notebooks', params: {}, expectedPath: '/me/onenote/notebooks' },
  { name: 'list-onenote-notebook-sections', params: { notebookId: 'n1' }, expectedPath: '/me/onenote/notebooks/n1/sections' },
  { name: 'list-all-onenote-sections', params: {}, expectedPath: '/me/onenote/sections' },
  { name: 'list-onenote-section-pages', params: { onenoteSectionId: 's1' }, expectedPath: '/me/onenote/sections/s1/pages' },
  { name: 'get-onenote-page-content', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'get-onenote-page-as-markdown', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'search-onenote-pages', params: { titleSubstring: 'meeting' }, expectedPath: "/me/onenote/pages?$filter=contains(title,'meeting')" },
  { name: 'get-current-user', params: {}, expectedPath: '/me' },
  { name: 'get-my-profile-photo', params: {}, expectedPath: '/me/photo/$value' },
  { name: 'list-calendar-events', params: {}, expectedPath: '/me/events' },
  { name: 'get-calendar-event', params: { eventId: 'e1' }, expectedPath: '/me/events/e1' },
  { name: 'list-specific-calendar-events', params: { calendarId: 'c1' }, expectedPath: '/me/calendars/c1/events' },
  { name: 'list-specific-calendar-events', params: { calendarId: 'primary' }, expectedPath: '/me/calendar/events' },
  { name: 'list-specific-calendar-events', params: { calendarId: 'DEFAULT' }, expectedPath: '/me/calendar/events' },
  { name: 'get-specific-calendar-event', params: { calendarId: 'c1', eventId: 'e1' }, expectedPath: '/me/calendars/c1/events/e1' },
  { name: 'get-specific-calendar-event', params: { calendarId: 'primary', eventId: 'e1' }, expectedPath: '/me/calendar/events/e1' },
  {
    name: 'list-calendar-view',
    params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendarView?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'list-specific-calendar-view',
    params: { calendarId: 'c1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendars/c1/calendarView?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'list-specific-calendar-view',
    params: { calendarId: 'primary', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendar/calendarView?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'list-calendar-event-instances',
    params: { calendarId: 'c1', eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendars/c1/events/e1/instances?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  {
    name: 'list-calendar-event-instances',
    params: { calendarId: 'primary', eventId: 'e1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendar/events/e1/instances?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  { name: 'list-calendars', params: {}, expectedPath: '/me/calendars' },
  { name: 'list-calendar-events-delta', params: {}, expectedPath: '/me/events/delta()' },
  {
    name: 'list-calendar-view-delta',
    params: { startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/me/calendarView/delta()?startDateTime=2026-04-01T00:00:00Z&endDateTime=2026-05-01T00:00:00Z',
  },
  { name: 'list-chat-members', params: { chatId: 'ch1' }, expectedPath: '/chats/ch1/members' },
  { name: 'list-joined-teams', params: {}, expectedPath: '/me/joinedTeams' },
  { name: 'get-team', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1' },
  { name: 'list-team-channels', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1/channels' },
  { name: 'get-team-channel', params: { teamId: 'tm1', channelId: 'ch1' }, expectedPath: '/teams/tm1/channels/ch1' },
  { name: 'list-chats', params: {}, expectedPath: '/me/chats' },
  { name: 'get-chat', params: { chatId: 'ch1' }, expectedPath: '/chats/ch1' },
  { name: 'list-my-direct-reports', params: {}, expectedPath: '/me/directReports' },
  { name: 'list-user-direct-reports', params: { userId: 'alice@contoso.com' }, expectedPath: '/users/alice@contoso.com/directReports' },
  { name: 'list-my-memberships', params: {}, expectedPath: '/me/memberOf' },
  { name: 'get-my-manager', params: {}, expectedPath: '/me/manager' },
  { name: 'get-user-manager', params: { userId: 'alice@contoso.com' }, expectedPath: '/users/alice@contoso.com/manager' },
  { name: 'list-relevant-people', params: {}, expectedPath: '/me/people' },
  { name: 'list-groups', params: {}, expectedPath: '/groups' },
  { name: 'get-group', params: { groupId: 'g1' }, expectedPath: '/groups/g1' },
  { name: 'list-group-members', params: { groupId: 'g1' }, expectedPath: '/groups/g1/members' },
  { name: 'list-group-owners', params: { groupId: 'g1' }, expectedPath: '/groups/g1/owners' },
  { name: 'list-group-events', params: { groupId: 'g1' }, expectedPath: '/groups/g1/events' },
  {
    name: 'list-group-calendar-view',
    params: { groupId: 'g1', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/groups/g1/calendarView?startDateTime=2026-04-01T00%3A00%3A00Z&endDateTime=2026-05-01T00%3A00%3A00Z',
  },
  { name: 'list-group-conversations', params: { groupId: 'g1' }, expectedPath: '/groups/g1/conversations' },
  { name: 'list-group-threads', params: { groupId: 'g1' }, expectedPath: '/groups/g1/threads' },
  { name: 'get-mail-message-mime', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1/$value' },
  { name: 'list-mail-folder-messages-delta', params: { mailFolderId: 'inbox' }, expectedPath: '/me/mailFolders/inbox/messages/delta()' },
  { name: 'list-shared-mailbox-messages', params: { userId: 'shared@contoso.com' }, expectedPath: '/users/shared@contoso.com/messages' },
  {
    name: 'list-shared-mailbox-folder-messages',
    params: { userId: 'shared@contoso.com', mailFolderId: 'inbox' },
    expectedPath: '/users/shared@contoso.com/mailFolders/inbox/messages',
  },
  { name: 'get-shared-mailbox-message', params: { userId: 'shared@contoso.com', messageId: 'm1' }, expectedPath: '/users/shared@contoso.com/messages/m1' },
  {
    name: 'list-conversation-messages',
    params: { conversationId: 'AAQkAD-conv-1' },
    expectedPath: "/me/messages?$filter=conversationId eq 'AAQkAD-conv-1'",
  },
  {
    name: 'list-conversation-messages',
    params: { conversationId: "AAQk'AD-conv-2" },
    expectedPath: "/me/messages?$filter=conversationId eq 'AAQk''AD-conv-2'",
  },
  { name: 'list-focused-inbox-overrides', params: {}, expectedPath: '/me/inferenceClassification/overrides' },
  { name: 'list-outlook-categories', params: {}, expectedPath: '/me/outlook/masterCategories' },
  { name: 'list-shared-calendar-events', params: { userId: 'colleague@contoso.com' }, expectedPath: '/users/colleague@contoso.com/calendar/events' },
  {
    name: 'list-shared-calendar-view',
    params: { userId: 'colleague@contoso.com', startDateTime: '2026-04-01T00:00:00Z', endDateTime: '2026-05-01T00:00:00Z' },
    expectedPath: '/users/colleague@contoso.com/calendarView?startDateTime=2026-04-01T00%3A00%3A00Z&endDateTime=2026-05-01T00%3A00%3A00Z',
  },
  { name: 'list-sharepoint-list-columns', params: { siteId: 's1', listId: 'l1' }, expectedPath: '/sites/s1/lists/l1/columns' },
  { name: 'get-sharepoint-list-column', params: { siteId: 's1', listId: 'l1', columnId: 'Title' }, expectedPath: '/sites/s1/lists/l1/columns/Title' },
  { name: 'list-sharepoint-site-onenote-notebooks', params: { siteId: 's1' }, expectedPath: '/sites/s1/onenote/notebooks' },
  { name: 'list-sharepoint-site-onenote-notebook-sections', params: { siteId: 's1', notebookId: 'nb1' }, expectedPath: '/sites/s1/onenote/notebooks/nb1/sections' },
  { name: 'list-sharepoint-site-onenote-section-pages', params: { siteId: 's1', onenoteSectionId: 'sec1' }, expectedPath: '/sites/s1/onenote/sections/sec1/pages' },
  { name: 'get-sharepoint-site-onenote-page-content', params: { siteId: 's1', onenotePageId: 'p1' }, expectedPath: '/sites/s1/onenote/pages/p1/content' },
  { name: 'list-drive-item-thumbnails', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/thumbnails' },
  { name: 'get-excel-used-range', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' }, expectedPath: '/drives/d1/items/i1/workbook/worksheets/Sheet1/usedRange()' },
  { name: 'list-rooms', params: {}, expectedPath: '/places/microsoft.graph.room' },
  { name: 'list-room-lists', params: {}, expectedPath: '/places/microsoft.graph.roomList' },
  { name: 'list-trending-insights', params: {}, expectedPath: '/me/insights/trending' },
  { name: 'list-recent-files', params: {}, expectedPath: '/me/drive/recent' },
  { name: 'list-shared-with-me', params: {}, expectedPath: '/me/drive/sharedWithMe' },
  { name: 'list-recently-used-insights', params: {}, expectedPath: '/me/insights/used' },
  { name: 'list-shared-insights', params: {}, expectedPath: '/me/insights/shared' },
  { name: 'get-organization', params: {}, expectedPath: '/organization' },
  { name: 'list-mail-folders-delta', params: {}, expectedPath: '/me/mailFolders/delta()' },
  { name: 'get-channel-files-folder', params: { teamId: 'tm1', channelId: 'ch1' }, expectedPath: '/teams/tm1/channels/ch1/filesFolder' },
  { name: 'get-drive-item-list-item', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/listItem' },
  { name: 'get-drive-item-analytics', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/analytics' },
  { name: 'list-team-installed-apps', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1/installedApps?$expand=teamsAppDefinition' },
  { name: 'list-calendar-groups', params: {}, expectedPath: '/me/calendarGroups' },
  { name: 'list-calendar-group-calendars', params: { calendarGroupId: 'cg1' }, expectedPath: '/me/calendarGroups/cg1/calendars' },
  { name: 'get-my-calendar', params: {}, expectedPath: '/me/calendar' },
  { name: 'list-site-columns', params: { siteId: 's1' }, expectedPath: '/sites/s1/columns' },
  { name: 'list-site-content-types', params: { siteId: 's1' }, expectedPath: '/sites/s1/contentTypes' },
  { name: 'list-sharepoint-site-pages', params: { siteId: 's1' }, expectedPath: '/sites/s1/pages' },
  { name: 'list-excel-defined-names', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/names' },
  { name: 'list-excel-worksheet-charts', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' }, expectedPath: '/drives/d1/items/i1/workbook/worksheets/Sheet1/charts' },
  { name: 'get-drive-special-folder', params: { folderName: 'documents' }, expectedPath: '/me/drive/special/documents' },
  { name: 'get-drive-root-delta', params: {}, expectedPath: '/me/drive/root/delta()' },
  { name: 'list-followed-drive-items', params: {}, expectedPath: '/me/drive/following' },
  { name: 'get-drive-item-created-by-user', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/createdByUser' },
  { name: 'get-drive-item-last-modified-by-user', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/lastModifiedByUser' },
  { name: 'get-site-analytics', params: { siteId: 's1' }, expectedPath: '/sites/s1/analytics' },
  { name: 'list-sharepoint-list-item-versions', params: { siteId: 's1', listId: 'l1', listItemId: '12' }, expectedPath: '/sites/s1/lists/l1/items/12/versions' },
  { name: 'get-mail-rule', params: { mailFolderId: 'inbox', messageRuleId: 'r1' }, expectedPath: '/me/mailFolders/inbox/messageRules/r1' },
  { name: 'list-excel-comments', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/workbook/comments' },
  {
    name: 'list-excel-worksheet-pivot-tables',
    params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' },
    expectedPath: '/drives/d1/items/i1/workbook/worksheets/Sheet1/pivotTables',
  },
  { name: 'list-sensitivity-labels', params: {}, expectedPath: '/me/informationProtection/sensitivityLabels' },
  { name: 'list-my-transitive-memberships', params: {}, expectedPath: '/me/transitiveMemberOf' },
  { name: 'get-team-primary-channel', params: { teamId: 'tm1' }, expectedPath: '/teams/tm1/primaryChannel' },
  { name: 'list-todo-tasks-delta', params: { todoTaskListId: 'tl1' }, expectedPath: '/me/todo/lists/tl1/tasks/delta()' },
  {
    name: 'next-page',
    params: { url: 'https://graph.microsoft.com/v1.0/me/messages?$skip=10' },
    expectedPath: '/me/messages?$skip=10',
  },
];

describe('all commands build correct Graph URL', () => {
  it.each(pathFixtures)('$name calls $expectedPath', async ({ name, params, expectedPath }) => {
    const url = await capturedUrl(name, params);
    expect(url).toBe(`https://graph.microsoft.com/v1.0${expectedPath}`);
  });
});

// Audit v1.0.0 §B1 — these collection endpoints all reject `$skip` server-side
// (`invalidRequest: $skip is not supported on this API.`). The CLI must NOT
// advertise `--skip` in the option set for these commands. Regression guard.
const NO_SKIP_COMMANDS: Array<{ name: string; meta: { options: ReadonlyArray<{ name: string }> } }> = [
  { name: 'list-folder-files', meta: listFolderFiles.meta },
  { name: 'list-drive-item-permissions', meta: listDriveItemPermissions.meta },
  { name: 'list-drive-item-versions', meta: listDriveItemVersions.meta },
  { name: 'list-drive-item-thumbnails', meta: listDriveItemThumbnails.meta },
  { name: 'search-onedrive-files', meta: searchOnedriveFiles.meta },
  { name: 'search-my-documents', meta: searchMyDocuments.meta },
  { name: 'get-drive-delta', meta: getDriveDelta.meta },
  { name: 'get-drive-root-delta', meta: getDriveRootDelta.meta },
  { name: 'list-recent-files', meta: listRecentFiles.meta },
  { name: 'list-followed-drive-items', meta: listFollowedDriveItems.meta },
  { name: 'list-sharepoint-site-drives', meta: listSharepointSiteDrives.meta },
  { name: 'list-sharepoint-site-list-items', meta: listSharepointSiteListItems.meta },
  { name: 'list-sharepoint-list-item-versions', meta: listSharepointListItemVersions.meta },
  { name: 'list-site-content-types', meta: listSiteContentTypes.meta },
  { name: 'list-sharepoint-site-pages', meta: listSharepointSitePages.meta },
  { name: 'list-groups', meta: listGroups.meta },
];

describe('no-skip endpoints do not advertise --skip', () => {
  it.each(NO_SKIP_COMMANDS)('$name.meta.options omits --skip', ({ meta }) => {
    const hasSkip = meta.options.some((o) => o.name === 'skip');
    expect(hasSkip).toBe(false);
  });
});

// Audit v1.0.0 §B1/B2/B3 — Graph's `/chats/{id}/members` rejects `$top`,
// `$orderby`, `$expand` with `BadRequest`; `/me/chats` rejects `$orderby`
// and hangs on `$expand`. Both endpoints DO honour the subset listed below.
// Regression guard: the CLI must advertise ONLY the working flags.
describe('chats endpoints advertise only the OData flags Graph honours', () => {
  it('list-chat-members.meta.options exposes skip/select/filter but NOT top/orderby/expand', () => {
    const names = listChatMembers.meta.options.map((o) => o.name).toSorted();
    expect(names).toEqual(['chat-id', 'filter', 'select', 'skip']);
  });

  it('list-chats.meta.options exposes top/skip/select/filter but NOT orderby/expand', () => {
    const names = listChats.meta.options.map((o) => o.name).toSorted();
    expect(names).toEqual(['filter', 'select', 'skip', 'top']);
  });

  it('list-chat-members drops $top/$orderby/$expand from the URL even when supplied as params', async () => {
    const url = await capturedUrl('list-chat-members', { chatId: 'ch1', top: '5', orderby: 'x', expand: 'members', select: 'id', filter: "x eq 'y'" });
    expect(url).toBe("https://graph.microsoft.com/v1.0/chats/ch1/members?$select=id&$filter=x%20eq%20'y'");
  });

  it('list-chats builds the URL with only the picked OData keys when params are valid', async () => {
    const url = await capturedUrl('list-chats', { top: '5', select: 'id' });
    expect(url).toBe('https://graph.microsoft.com/v1.0/me/chats?$top=5&$select=id');
  });
});

// Audit v1.0.0 §B9 — the single-resource GET on a Microsoft task list was
// the only "get" without `--select`. Sister GETs (get-my-manager,
// get-user-manager, get-mail-message, etc.) all expose `--select`/`--expand`
// so an LLM can slim the response payload. Regression guard: advertise both.
describe('get-todo-task supports --select and --expand', () => {
  it("the command's meta.options includes 'select' and 'expand' so an LLM can slim the response payload", () => {
    const names = getTodoTask.meta.options.map((o) => o.name);
    expect(names).toContain('select');
    expect(names).toContain('expand');
  });

  it('the command forwards --select to Graph as $select', async () => {
    const url = await capturedUrl('get-todo-task', { todoTaskListId: 'tl1', todoTaskId: 't1', select: 'id,title' });
    expect(url).toBe('https://graph.microsoft.com/v1.0/me/todo/lists/tl1/tasks/t1?$select=id%2Ctitle');
  });
});

// Audit v1.0.0 §B6 / D1 — Graph rejects `$search` with `$filter` together
// (the real Graph error is `SearchWithFilter`, not the docs' previous
// `InvalidRestriction`). Reject the conflict client-side before round-trip
// so the LLM gets a precise pointer to the right alternative command
// instead of an opaque Graph code.
describe('search-mail-messages rejects --filter client-side', () => {
  it('returns validation_error when --filter is supplied alongside --query (Graph does not allow $search + $filter together)', async () => {
    const cmd = cmdMap['search-mail-messages'];
    if (!cmd) throw new Error('search-mail-messages not registered');
    const fetchFn = fakeFetch({ value: [] });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { query: 'invoice', filter: "from/emailAddress/address eq 'alice'" });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).toContain('--filter');
      expect(result.error.message).toContain('list-mail-messages');
    }
    expect(fetchFn.lastUrl).toBeNull();
  });

  it('still works when only --query is supplied (no regression on the happy path)', async () => {
    const url = await capturedUrl('search-mail-messages', { query: 'invoice' });
    expect(url).toBe('https://graph.microsoft.com/v1.0/me/messages?$search="invoice"');
  });

  it("on an empty `--filter ''` the client-side guard does NOT fire — the rejection guard checks for a non-empty filter, so the message comes from the downstream Zod min-1 check, not the `$search`-conflict pointer (boundary on params['filter'].length > 0)", async () => {
    const fetchFn = fakeFetch({ value: [] });
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await searchMailMessages.execute(graph, { query: 'invoice', filter: '' });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).not.toContain('incompatible with $search');
    }
  });

  it('renders docs for search-mail-messages with the documented response shape, the KQL-query option description, and the pagination hint — guards the meta block content against silent rewrites', () => {
    const rendered = renderSingleCommand(cmdRegistry, 'search-mail-messages');
    expect(rendered.ok).toBe(true);
    if (rendered.ok) {
      expect(rendered.value).toContain('GET /me/messages');
      expect(rendered.value).toContain('ranked by relevance');
      expect(rendered.value).toContain('KQL or free-text query');
      expect(rendered.value).toContain('Examples: ');
      expect(rendered.value.toLowerCase()).toContain('pagination');
    }
  });
});
