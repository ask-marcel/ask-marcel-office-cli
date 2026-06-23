import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';
import type { FetchFn, GraphError } from '../../infra/graph-client.ts';
import { createGraphClient } from '../../infra/graph-client.ts';
import {
  buildDocxWithSharepointLinks,
  buildMalformedDocx,
  buildMediaSamples,
  buildOdtWithSharepointLinks,
  buildOversizedZipArchive,
  buildPdfWithImage,
  buildPdfWithText,
  buildRichOdt,
  buildRichPptx,
  buildSampleDocx,
  buildSampleMsg,
  buildSampleXlsx,
  buildSampleZipArchive,
} from '../../test-helpers/office-fixtures.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { renderSingleCommand } from './docs.ts';
import { commands as cmdRegistry } from './index.ts';
import * as downloadDriveItemAsMarkdown from './download-drive-item-as-markdown.ts';
import * as extractDriveItemImages from './extract-drive-item-images.ts';
import * as listAccessibleDrives from './list-accessible-drives.ts';
import * as downloadDriveItemAsPdf from './download-drive-item-as-pdf.ts';
import * as downloadDriveItemVersion from './download-drive-item-version.ts';
import * as downloadDriveItemContent from './download-drive-item-content.ts';
import * as getCalendarEvent from './get-calendar-event.ts';
import * as getCalendarView from './get-calendar-view.ts';
import * as getCurrentUser from './get-current-user.ts';
import * as getDriveDelta from './get-drive-delta.ts';
import * as getDriveItem from './get-drive-item.ts';
import * as getDriveRootItem from './get-drive-root-item.ts';
import * as getExcelChartImage from './get-excel-chart-image.ts';
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
import * as searchAllAccessibleSites from './search-all-accessible-sites.ts';
import * as searchSharepointSitesByName from './search-sharepoint-sites-by-name.ts';
import * as extractSharepointLinksInMail from './extract-sharepoint-links-in-mail.ts';
import * as convertMailAttachmentToMarkdown from './convert-mail-attachment-to-markdown.ts';
import * as extractMailAttachmentImages from './extract-mail-attachment-images.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as listCalendarEventAttachments from './list-calendar-event-attachments.ts';
import * as convertCalendarEventAttachmentToMarkdown from './convert-calendar-event-attachment-to-markdown.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';
import * as convertMailToMarkdown from './convert-mail-to-markdown.ts';
import * as convertDriveItemZip from './convert-drive-item-zip.ts';
import * as extractSharepointLinksInDocuments from './extract-sharepoint-links-in-documents.ts';
import { buildShareToken, resolveSharepointUrls } from './sharepoint-link-extractor.ts';
import { stripQuotedPlainText, stripQuotedReplies } from './mail-quote-stripper.ts';
import { embedOnenoteResources } from './onenote-resource-embedder.ts';
import { formatOnenoteMetadata } from './onenote-metadata.ts';
import * as listChats from './list-chats.ts';
import * as getChat from './get-chat.ts';
import * as listTeamsChatsWithMessages from './list-teams-chats-with-messages.ts';
import * as listTeamsChatMessages from './list-teams-chat-messages.ts';
import * as listTeamsChatHistory from './list-teams-chat-history.ts';
import * as resolveTeamsLink from './resolve-teams-link.ts';
import * as resolveMailLink from './resolve-mail-link.ts';
import * as resolveDriveShareLink from './resolve-drive-share-link.ts';
import * as resolveCalendarLink from './resolve-calendar-link.ts';
import * as findChatsWithUser from './find-chats-with-user.ts';
import * as getTeamsChatMessage from './get-teams-chat-message.ts';
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
  'download-drive-item-content': downloadDriveItemContent,
  'get-drive-item': getDriveItem,
  'list-drive-item-permissions': listDriveItemPermissions,
  'list-drive-item-versions': listDriveItemVersions,
  'download-drive-item-version': downloadDriveItemVersion,
  'download-drive-item-as-pdf': downloadDriveItemAsPdf,
  'download-drive-item-as-markdown': downloadDriveItemAsMarkdown,
  'extract-drive-item-images': extractDriveItemImages,
  'list-accessible-drives': listAccessibleDrives,
  'search-onedrive-files': searchOnedriveFiles,
  'search-my-documents': searchMyDocuments,
  'get-excel-chart-image': getExcelChartImage,
  'get-excel-range': getExcelRange,
  'list-excel-worksheets': listExcelWorksheets,
  'list-excel-tables': listExcelTables,
  'get-excel-table': getExcelTable,
  'list-excel-table-rows': listExcelTableRows,
  'get-drive-delta': getDriveDelta,
  'search-sharepoint-sites-by-name': searchSharepointSitesByName,
  'search-all-accessible-sites': searchAllAccessibleSites,
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
  'convert-drive-item-zip': convertDriveItemZip,
  'extract-sharepoint-links-in-documents': extractSharepointLinksInDocuments,
  'convert-mail-attachment-to-pdf': convertMailAttachmentToPdf,
  'convert-mail-attachment-to-markdown': convertMailAttachmentToMarkdown,
  'list-calendar-event-attachments': listCalendarEventAttachments,
  'convert-calendar-event-attachment-to-markdown': convertCalendarEventAttachmentToMarkdown,
  'convert-calendar-event-attachment-to-pdf': convertCalendarEventAttachmentToPdf,
  'extract-mail-attachment-images': extractMailAttachmentImages,
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
  'list-teams-chats-with-messages': listTeamsChatsWithMessages,
  'list-teams-chat-messages': listTeamsChatMessages,
  'list-teams-chat-history': listTeamsChatHistory,
  'get-teams-chat-message': getTeamsChatMessage,
  'resolve-teams-link': resolveTeamsLink,
  'resolve-mail-link': resolveMailLink,
  'resolve-drive-share-link': resolveDriveShareLink,
  'resolve-calendar-link': resolveCalendarLink,
  'find-chats-with-user': findChatsWithUser,
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
  getChatsvcaggAccessToken: async () => ok(accessTokenUnsafe('test-chatsvcagg-token')),
  getChatsvcaggRegion: async () => 'emea',
  getIc3AccessToken: async () => ok(accessTokenUnsafe('test-ic3-token')),
  getLastChatsvcaggOutcome: () => null,
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

  // Regression: a user report flagged that `list-chats` page 1 returned
  // `data: { value: [...] }` while page 2 via `next-page` returned
  // `data: [...]` directly. The static read shows both calls route
  // through `graph.getElevated()` and hit the same `wrap()`, so envelope
  // shapes SHOULD match. Pin that with a mocked test so any future
  // divergence (e.g. someone adds a `next-page`-only response transform)
  // fails loudly.
  it('list-chats and next-page produce structurally identical response values for the same Graph response', async () => {
    const sameBody = { '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#chats', value: [{ id: '19:abc', topic: 'A' }] };
    const list = await callCommand('list-chats', {}, sameBody);
    const next = await callCommand('next-page', { url: 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=opaque' }, sameBody);
    expect(list.ok).toBe(true);
    expect(next.ok).toBe(true);
    if (list.ok && next.ok) {
      // Both must hand the SAME object back to the presenter — wrap() then
      // hoists `@odata.nextLink`/`deltaLink`/`count` identically.
      expect(list.value).toEqual(next.value);
    }
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

  it('get-excel-chart-image returns the chart PNG as { image/png, size, base64 } from Graph chart Image()', async () => {
    const result = await callCommand('get-excel-chart-image', { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', chartId: 'Chart 1' }, { value: 'aW1n' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType: string; size: number; base64: string };
    expect(v.contentType).toBe('image/png');
    expect(v.base64).toBe('aW1n');
    expect(v.size).toBe(3);
  });

  it('get-excel-chart-image errs when Graph returns no base64 string (wrong type OR empty string) in the value field', async () => {
    for (const value of [42, '']) {
      const result = await callCommand('get-excel-chart-image', { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', chartId: 'c1' }, { value });
      expect(result.ok).toBe(false);
      if (result.ok) continue;
      expect(result.error.type).toBe('api_error');
      expect(result.error.message).toContain('no base64 PNG');
    }
  });

  it('get-excel-chart-image returns a validation_error when chartId is missing', async () => {
    const result = await callCommand('get-excel-chart-image', { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' }, { value: 'x' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('get-excel-chart-image maps the non-workbook Graph error through the excel-error wrapper', async () => {
    const graph = createGraphClient(
      fakeAuth(),
      async () =>
        new Response(JSON.stringify({ error: { code: 'AccessDenied', message: 'Could not obtain a WAC access token' } }), {
          status: 403,
          headers: { 'content-type': 'application/json' },
        })
    );
    const cmd = cmdMap['get-excel-chart-image'];
    if (!cmd) throw new Error('get-excel-chart-image not registered');
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', chartId: 'c1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect((result.error as { message: string }).message).toContain('not an accessible Excel workbook');
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

  it('download-drive-item-content returns binary bytes as faithful inline base64 (content-sniffed — not coerced to text)', async () => {
    const binaryB64 = btoa(String.fromCharCode(0xff, 0xfe, 0xfd, 0x80, 0x81)); // invalid UTF-8 → sniffs binary
    const result = await callCommand('download-drive-item-content', { driveId: 'd1', itemId: 'i1' }, { contentType: 'application/octet-stream', size: 5, base64: binaryB64 });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/octet-stream');
      expect(v.base64).toBe(binaryB64); // bytes returned intact, never mangled into `�`
    }
  });

  it('download-drive-item-content returns valid-UTF-8 bytes as text (content-sniffed, regardless of extension)', async () => {
    const utf8 = new TextEncoder().encode('hi 文'); // multi-byte: proves the sniffer decodes real UTF-8, not just ASCII
    const textB64 = btoa(String.fromCharCode(...utf8));
    const result = await callCommand(
      'download-drive-item-content',
      { driveId: 'd1', itemId: 'i1' },
      { contentType: 'application/octet-stream', size: utf8.byteLength, base64: textB64 }
    );
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toBe('hi 文');
    }
  });

  it('download-drive-item-content rejects a folder --item-id with a clear "this is a folder, use list-folder-files" hint instead of empty error (audit round-6 §1.1)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFolder', method: 'GET', response: Response.json({ name: 'Reports', folder: { childCount: 12 } }) },
    ]);
    const cmd = cmdMap['download-drive-item-content'];
    if (!cmd) throw new Error('download-drive-item-content not registered');
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
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toContain('Graph rejected --orderby=title asc');
    expect(r.error.message).toContain('sorting on `title` is unsupported');
    expect(r.error.code).toBe('cli_rewrite_todo_orderby_title');
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

  it('get-todo-task rewrites the opaque RequestBroker--ParseUri error to the title-quirk hint when --select includes title', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1?$select=id%2Ctitle',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['get-todo-task'];
    if (!cmd) throw new Error('get-todo-task not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', todoTaskId: 't1', select: 'id,title' });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toContain('Graph rejected --select=id,title');
    expect(r.error.message).toContain('Drop `title` from --select');
    expect(r.error.message).toContain('slim the response client-side');
    expect(r.error.message).not.toContain('per-task');
    expect(r.error.code).toBe('cli_rewrite_todo_select_title');
  });

  it('get-todo-task passes through RequestBroker--ParseUri unchanged when no --select is set (no false rewrite)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'RequestBroker--ParseUri', message: 'Invalid request' } }), {
            status: 400,
            headers: { 'content-type': 'application/json' },
          }),
      },
    ]);
    const cmd = cmdMap['get-todo-task'];
    if (!cmd) throw new Error('get-todo-task not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', todoTaskId: 't1' });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toContain('RequestBroker--ParseUri');
    expect(r.error.message).not.toContain('Graph rejected');
    expect(r.error.message).not.toContain('Drop `title`');
  });

  it('get-todo-task passes through a non-ParseUri error unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { code: 'Unauthorized', message: 'bad token' } }), { status: 401, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['get-todo-task'];
    if (!cmd) throw new Error('get-todo-task not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', todoTaskId: 't1' });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toBe('Unauthorized: bad token');
  });

  it('get-todo-task does NOT rewrite a non-ParseUri error even when --select is set (only the title quirk is special-cased)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/todo/lists/l1/tasks/t1?$select=id%2Ctitle',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { code: 'Unauthorized', message: 'bad token' } }), { status: 401, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['get-todo-task'];
    if (!cmd) throw new Error('get-todo-task not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { todoTaskListId: 'l1', todoTaskId: 't1', select: 'id,title' });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toBe('Unauthorized: bad token');
    expect(r.error.message).not.toContain('Drop `title`');
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

  // Regression guard for the 2026-06 elevation drop: `/chats/{id}/members` is
  // served by the BASIC Teams token (`ChatMember.Read`), NOT the login-only
  // elevated M365 token. The path-fixture / error-rewrite tests are
  // token-agnostic (real client + fake fetch matched on URL), so only this
  // tracking fake catches a silent revert to `graph.getElevated` — which would
  // resurrect the "Elevated token expired, run login" failure the switch fixed.
  it('list-chat-members reads /chats/{id}/members via the basic token (graph.get), not the elevated M365 token', async () => {
    const via: Array<'basic' | 'elevated'> = [];
    const graph = fakeGraphClient({
      get: async () => {
        via.push('basic');
        return ok({ value: [] });
      },
      getElevated: async () => {
        via.push('elevated');
        return ok({ value: [] });
      },
    });
    const cmd = cmdMap['list-chat-members'];
    if (!cmd) throw new Error('list-chat-members not registered');
    const r = await cmd.execute(graph, { chatId: '19:abc@thread.v2' });
    expect(r.ok).toBe(true);
    expect(via).toEqual(['basic']);
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
    if (r.ok) return;
    expect(r.error.type).toBe('api_error');
    if (r.error.type !== 'api_error') return;
    expect(r.error.message).toContain('Graph rejected --orderby=title asc');
    expect(r.error.message).toContain('sorting on `title` is unsupported');
    expect(r.error.code).toBe('cli_rewrite_todo_orderby_title');
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
    if (r.ok) return;
    expect(r.error.type).toBe('validation_error');
    if (r.error.type !== 'validation_error') return;
    expect(r.error.message).toContain('--filter is not supported on list-incomplete-todo-tasks');
    expect(r.error.message).toContain('list-todo-tasks');
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

  it('download-drive-item-content returns plain-text source extensions inline as `{contentType: "text/plain", text}` (no 33% base64 bloat — audit v1.0.0 §bug-3)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'README.md', size: 5 }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/content',
        method: 'GET',
        response: () => new Response(new TextEncoder().encode('# hi'), { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-content'];
    if (!cmd) throw new Error('download-drive-item-content not registered');
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

  it('download-drive-item-version (default --format original) inlines the historical-version bytes via the M365ChatClient-elevated path', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content',
        method: 'GET',
        response: () => Response.json({ '@microsoft.graph.downloadUrl': 'https://contoso.sharepoint.com/cdn/v3.bin' }),
      },
      {
        urlPrefix: 'https://contoso.sharepoint.com/cdn/v3.bin',
        method: 'GET',
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
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
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } }),
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

  it('download-drive-item-as-pdf short-circuits plain-text source extensions to `{contentType: "text/plain", size, text}` for envelope parity with download-drive-item-content (audit round-7 B5)', async () => {
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

  it('download-drive-item-as-markdown renders an Outlook .msg with its attachment recursed inline', async () => {
    const msgBytes = await buildSampleMsg();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMsg', method: 'GET', response: Response.json({ name: 'email.msg' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMsg/content',
        method: 'GET',
        response: () => new Response(msgBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-as-markdown'];
    if (!cmd) throw new Error('download-drive-item-as-markdown not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iMsg' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType: string; text: string };
    expect(v.contentType).toBe('text/markdown');
    expect(v.text).toContain('# Quarterly Report — Q3 Summary');
    expect(v.text).toContain('**From:** Jordan Avery <jordan.avery@example.com>');
    expect(v.text).toContain('## Attachments');
    expect(v.text).toContain('### summary.txt');
    expect(v.text).toContain('Attachment body text'); // the .txt attachment recursed to text/plain
  });

  it('extract-drive-item-images returns a base64 media envelope of the images (raster + svg) embedded in a pptx', async () => {
    const pptxBytes = await buildMediaSamples();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iDeck', method: 'GET', response: Response.json({ name: 'deck.pptx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iDeck/content',
        method: 'GET',
        response: () => new Response(pptxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iDeck' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { count: number; media: ReadonlyArray<{ path: string; contentType: string; sizeBytes: number; base64: string }> };
      expect(v.count).toBe(4);
      expect(v.media.map((m) => m.path)).toEqual(['ppt/media/diagram.gif', 'word/media/chart.svg', 'word/media/image1.png', 'xl/media/photo.jpeg']);
      const png = v.media.find((m) => m.path === 'word/media/image1.png');
      expect(png?.contentType).toBe('image/png');
      expect(png?.sizeBytes).toBe(4);
      expect(typeof png?.base64).toBe('string');
      expect(v.media.find((m) => m.path === 'word/media/chart.svg')?.contentType).toBe('image/svg+xml');
    }
  });

  it('extract-drive-item-images extracts PNG-encoded page images from a PDF source', async () => {
    const pdfBytes = buildPdfWithImage();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf', method: 'GET', response: Response.json({ name: 'report.pdf' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iPdf/content',
        method: 'GET',
        response: () => new Response(pdfBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iPdf' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; media: ReadonlyArray<{ path: string; contentType: string }> };
    expect(v.count).toBe(1);
    expect(v.media[0]?.path).toMatch(/^pdf\/page1\/.+\.png$/);
    expect(v.media[0]?.contentType).toBe('image/png');
  });

  it('extract-drive-item-images rejects an unsupported source with a 415 that names the extension and points at download-drive-item-content', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iTxt/content', method: 'GET', response: () => new Response(new Uint8Array([1, 2, 3]), { status: 200 }) },
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iTxt', method: 'GET', response: Response.json({ name: 'notes.txt' }) },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iTxt' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.status).toBe(415);
    expect(result.error.message).toContain('txt is not a supported document — image extraction supports pdf and docx / xlsx / pptx');
    expect(result.error.message).toContain('download-drive-item-content');
  });

  it('extract-drive-item-images returns a validation_error when itemId is missing', async () => {
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), stagedFetch([])), { driveId: 'd1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });

  it('extract-drive-item-images surfaces a media-extraction failure when an OOXML-named file is not a valid zip', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iBad', method: 'GET', response: Response.json({ name: 'corrupt.docx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iBad/content',
        method: 'GET',
        response: () => new Response(buildMalformedDocx() as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iBad' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('ooxml media extraction failed');
  });

  it('extract-drive-item-images treats a trailing-dot name (no real extension) as non-OOXML', async () => {
    const fetchFn = stagedFetch([{ urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iDot', method: 'GET', response: Response.json({ name: 'weird.' }) }]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iDot' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('<no-extension> is not a supported document');
  });

  it('extract-drive-item-images aliases the docx and xlsx macro-enabled families onto the same extractor', async () => {
    const bytes = await buildMediaSamples();
    for (const [name, itemId] of [
      ['macro.docm', 'iDocm'],
      ['model.xlsm', 'iXlsm'],
    ] as const) {
      const fetchFn = stagedFetch([
        { urlPrefix: `https://graph.microsoft.com/v1.0/drives/d1/items/${itemId}`, method: 'GET', response: Response.json({ name }) },
        {
          urlPrefix: `https://graph.microsoft.com/v1.0/drives/d1/items/${itemId}/content`,
          method: 'GET',
          response: () => new Response(bytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
        },
      ]);
      const cmd = cmdMap['extract-drive-item-images'];
      if (!cmd) throw new Error('extract-drive-item-images not registered');
      const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId });
      expect(result.ok).toBe(true);
      if (result.ok) expect((result.value as { count: number }).count).toBe(4);
    }
  });

  it('extract-drive-item-images errs 400 with a list-folder-files hint when the item is a folder (mirrors download-drive-item-as-pdf)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFolder', method: 'GET', response: Response.json({ name: 'My Folder', folder: { childCount: 3 } }) },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iFolder' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(400);
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain("'My Folder' is a folder");
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('list-folder-files');
  });

  it('extract-drive-item-images returns a 415 api_error with the <no-extension> placeholder when the driveItem name has no extension', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iNoExt/content', method: 'GET', response: () => new Response(new Uint8Array([1, 2, 3]), { status: 200 }) },
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iNoExt', method: 'GET', response: Response.json({ name: 'README' }) },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iNoExt' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.status).toBe(415);
    expect(result.error.message).toContain('<no-extension> is not a supported document');
  });

  it('extract-drive-item-images returns count 0 with an empty media array for an OOXML doc with no embedded images', async () => {
    const xlsxBytes = buildSampleXlsx();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iEmpty', method: 'GET', response: Response.json({ name: 'plain.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iEmpty/content',
        method: 'GET',
        response: () => new Response(xlsxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iEmpty' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { count: number; media: ReadonlyArray<unknown> };
      expect(v.count).toBe(0);
      expect(v.media).toEqual([]);
    }
  });

  it('extract-drive-item-images propagates the driveItem metadata GET error (404) verbatim, before any extension check', async () => {
    const fetchFn = stagedFetch([{ urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iGone', method: 'GET', response: new Response('nope', { status: 404 }) }]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iGone' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(404);
  });

  it('extract-drive-item-images propagates a content-fetch failure verbatim (503), not the extraction error path', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFail', method: 'GET', response: Response.json({ name: 'deck.pptx' }) },
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iFail/content', method: 'GET', response: new Response('boom', { status: 503 }) },
    ]);
    const cmd = cmdMap['extract-drive-item-images'];
    if (!cmd) throw new Error('extract-drive-item-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'iFail' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(503);
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

  it('download-drive-item-version --format pdf converts a non-current version through Graph ?format=pdf and inlines the PDF bytes (CDN redirect followed via the M365ChatClient-elevated path)', async () => {
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
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0', format: 'pdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-version --format markdown converts a non-current xlsx version via the local sheetjs pipeline', async () => {
    const xlsxBytes = buildSampleXlsx();
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1', method: 'GET', response: Response.json({ name: 'budget.xlsx' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/i1/versions/3.0/content',
        method: 'GET',
        response: () => new Response(xlsxBytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', versionId: '3.0', format: 'markdown' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('## Sheet1');
    }
  });

  it('download-drive-item-version --format markdown short-circuits to raw download for plain-text source extensions', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'notes.md' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('# v2', { status: 200, headers: { 'content-type': 'text/markdown' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0', format: 'markdown' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toBe('# v2');
    }
  });

  it('download-drive-item-version --format pdf short-circuits to a raw bytes download for plain-text source extensions (re-encodes text body as base64)', async () => {
    const fetchFn = stagedFetch([
      { urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText', method: 'GET', response: Response.json({ name: 'log.log' }) },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iText/versions/2.0/content',
        method: 'GET',
        response: () => new Response('line', { status: 200, headers: { 'content-type': 'text/plain' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iText', versionId: '2.0', format: 'pdf' });
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
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } }),
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

  it('download-drive-item-version --format pdf short-circuits to raw download for a pdf source (same reason as the non-versioned variant) and inlines the bytes', async () => {
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
        response: () => new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iPdf', versionId: '2.0', format: 'pdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; base64: string };
      expect(v.contentType).toBe('application/pdf');
      expect(atob(v.base64)).toBe('%PDF-');
    }
  });

  it('download-drive-item-version --format pdf propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0', format: 'pdf' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
    }
  });

  it('download-drive-item-version --format markdown propagates an err from the metadata pre-fetch unchanged', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/drives/d1/items/iMissing',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['download-drive-item-version'];
    if (!cmd) throw new Error('download-drive-item-version not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'iMissing', versionId: '2.0', format: 'markdown' });
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

  it('convert-mail-to-markdown renders a single email with headers + body via turndown and skips the attachments-list call entirely when hasAttachments is false (one round-trip)', async () => {
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
            hasAttachments: false,
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
      expect(v.text).not.toContain('**Attachments:**');
      // Header lines are NEWLINE-separated, not concatenated (kills renderHeaders join('\n') → join('') mutant).
      expect(v.text).toContain('**Subject:** Q3\n**From:**');
      // Body section is preceded by exactly ONE blank line, not zero — kills the L169 '\n\n' → 'Stryker' join-separator mutant.
      expect(v.text).toContain('**Date:** 2026-04-30T08:00:00Z\n\nHi');
      // Text doesn't start with a blank line (kills L169 `s !== ''` → true mutant — the filter must actually drop empty strings, otherwise an empty fileList becomes a leading '\n\n').
      expect(v.text.startsWith('**Subject:**')).toBe(true);
    }
  });

  it('convert-mail-to-markdown fetches per-attachment bytes lazily for inline image-only — heavy file attachments are listed by name+size and their bytes never round-trip (audit v1.0.0 multi-MB-attachment timeout fix)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m2',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'with logo and heavy attachment',
            body: { contentType: 'html', content: '<p>Logo: <img src="cid:logo123" alt="logo"></p>' },
            hasAttachments: true,
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m2/attachments?',
        method: 'GET',
        response: () =>
          Response.json({
            value: [
              { id: 'a1', name: 'logo.png', contentType: 'image/png', contentId: 'logo123', isInline: true, size: 256 },
              { id: 'a2', name: 'report.pdf', contentType: 'application/pdf', isInline: false, size: 4_200_000 },
            ],
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m2/attachments/a1',
        method: 'GET',
        response: () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', contentBytes: 'iVBORw0=' }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm2' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      // Inline image embedded via the per-attachment bytes fetch.
      expect(v.text).toContain('data:image/png;base64,iVBORw0=');
      expect(v.text).not.toContain('cid:logo123');
      // Heavy file attachment listed by metadata only — no bytes ever fetched.
      expect(v.text).toContain('**Attachments:**');
      expect(v.text).toContain('report.pdf');
      expect(v.text).toContain('4.2 MB');
      expect(v.text).toContain('a2');
    }
  });

  it('convert-mail-to-markdown skips per-attachment bytes-fetch for non-image inline attachments and never embeds their contentBytes (Hardening #1)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m3',
        method: 'GET',
        response: () => Response.json({ subject: 'sneaky', body: { contentType: 'html', content: '<p>X: <img src="cid:evil"></p>' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m3/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [{ id: 'aEvil', name: 'evil.html', contentType: 'text/html', contentId: 'evil', isInline: true, size: 128 }] }),
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

  it('convert-mail-to-markdown skips per-attachment bytes-fetch for inline images larger than the 2 MB guard and emits a placeholder in the markdown', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mHuge',
        method: 'GET',
        response: () => Response.json({ subject: 'huge inline', body: { contentType: 'html', content: '<p>X: <img src="cid:huge"></p>' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mHuge/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [{ id: 'aHuge', name: 'huge.png', contentType: 'image/png', contentId: 'huge', isInline: true, size: 5_200_000 }] }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mHuge' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      // Specific placeholder format — kills string-literal mutants on the label prefix/separator.
      // Parens get escaped by turndown when the placeholder lands inside an HTML <img> alt-text path,
      // so assert the unescaped fragments individually.
      expect(v.text).toContain('inline image too large to embed: huge.png');
      expect(v.text).toContain('5.2 MB');
      expect(v.text).not.toContain('cid:huge');
      expect(v.text).not.toContain('data:image/png');
    }
  });

  it('convert-mail-to-markdown returns validation_error when the attachments-list returns a malformed shape (Graph glitch returning value: "not an array") — boundary validation prevents the downstream TypeError on .filter()', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBadShape',
        method: 'GET',
        response: () => Response.json({ subject: 'bad-shape', body: { contentType: 'text', content: 'b' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBadShape/attachments?',
        method: 'GET',
        response: () => Response.json({ value: 'not-an-array' }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mBadShape' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      // Body in hand is still rendered; the malformed metadata is surfaced as a note (failure isolation).
      const v = result.value as { text: string; note?: string };
      expect(v.text).toContain('bad-shape');
      expect(v.note).toBeDefined();
      expect(v.note ?? '').toContain('attachments-list');
    }
  });

  it('convert-mail-to-markdown returns the markdown body with a note when the attachments-list fetch fails — the body in hand is still useful, the partial-success preserves it (failure isolation)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mPartial',
        method: 'GET',
        response: () => Response.json({ subject: 'partial', body: { contentType: 'html', content: '<p>Body here.</p>' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mPartial/attachments?',
        method: 'GET',
        response: () =>
          new Response(JSON.stringify({ error: { code: 'ServiceUnavailable', message: 'partial' } }), { status: 503, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mPartial' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string; note?: string };
      expect(v.text).toContain('Body here.');
      expect(v.note).toBeDefined();
      expect(v.note ?? '').toContain('attachment');
    }
  });

  it('convert-mail-to-markdown handles plain-text bodies WITHOUT turndown processing — HTML-looking markup is preserved literally when contentType is text (kills the `m.body?.contentType === "html"` → true mutant which would otherwise strip the tags)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/m4',
        method: 'GET',
        response: () => Response.json({ subject: 'plain', body: { contentType: 'text', content: '<p>Hello</p>\n<br>world' }, hasAttachments: false }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm4' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      // The literal HTML tags survive — turndown was NOT called.
      expect(v.text).toContain('<p>Hello</p>');
      expect(v.text).toContain('<br>world');
      // Text does not end with extra blank lines — kills the L169 .filter() mutant
      // (without filter, an empty fileList would tail '\n\n' onto the join).
      expect(v.text.endsWith('\n\n')).toBe(false);
      expect(v.text.endsWith('\n')).toBe(false);
    }
  });

  it('convert-mail-to-markdown defaults body.content to empty string when missing — no `Stryker was here!` or other garbage leaks into the output (kills the L157 StringLiteral `?? ""` → `?? "Stryker"` mutant)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNoBody',
        method: 'GET',
        response: () => Response.json({ subject: 'no-body', body: { contentType: 'text' }, hasAttachments: false }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mNoBody' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).not.toContain('Stryker');
      // The output is just the headers — kills the L141 ArrayDeclaration mutant
      // that would have started attachments with a sentinel string element.
      expect(v.text).toBe('**Subject:** no-body');
    }
  });

  it('convert-mail-to-markdown embeds a 2 MB inline image at the exact boundary (size === 2_000_000) — the size guard rejects only sizes STRICTLY greater than 2 MB, not equal', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBoundary',
        method: 'GET',
        response: () => Response.json({ subject: 'boundary', body: { contentType: 'html', content: '<p>I: <img src="cid:edge"></p>' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBoundary/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [{ id: 'aEdge', name: 'edge.png', contentType: 'image/png', contentId: 'edge', isInline: true, size: 2_000_000 }] }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mBoundary/attachments/aEdge',
        method: 'GET',
        response: () => Response.json({ contentBytes: 'AAAA' }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mBoundary' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('data:image/png;base64,AAAA');
      expect(v.text).not.toContain('inline image too large to embed');
    }
  });

  it('convert-mail-to-markdown fetches multiple inline images in parallel and leaves the cid: ref in place when an individual per-attachment fetch fails (failure isolation per-image)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mMulti',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'multi',
            body: { contentType: 'html', content: '<p>A:<img src="cid:a"> B:<img src="cid:b"> C:<img src="cid:c"></p>' },
            hasAttachments: true,
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mMulti/attachments?',
        method: 'GET',
        response: () =>
          Response.json({
            value: [
              { id: 'a1', name: 'a.png', contentType: 'image/png', contentId: 'a', isInline: true, size: 256 },
              { id: 'a2', name: 'b.jpg', contentType: 'image/jpeg', contentId: 'b', isInline: true, size: 256 },
              { id: 'a3', name: 'c.png', contentType: 'image/png', contentId: 'c', isInline: true, size: 256 },
            ],
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mMulti/attachments/a1',
        method: 'GET',
        response: () => Response.json({ contentBytes: 'BBBB' }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mMulti/attachments/a2',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { code: 'NotFound', message: 'lost' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mMulti/attachments/a3',
        method: 'GET',
        response: () => Response.json({ contentBytes: '' }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mMulti' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('data:image/png;base64,BBBB');
      expect(v.text).toContain('cid:b');
      expect(v.text).toContain('cid:c');
    }
  });

  it('convert-mail-to-markdown isInlineImage predicate rejects: non-image contentType (application/pdf), missing contentId, empty contentId, isInline:false with image contentType, isInline:undefined — all five fall through to the file-list path instead of being embed candidates', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mPred',
        method: 'GET',
        response: () =>
          Response.json({
            subject: 'predicate',
            body: { contentType: 'html', content: '<p>X: <img src="cid:r1"><img src="cid:r2"><img src="cid:r4"><img src="cid:r5"></p>' },
            hasAttachments: true,
          }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mPred/attachments?',
        method: 'GET',
        response: () =>
          Response.json({
            value: [
              { id: 'r1', name: 'r1.pdf', contentType: 'application/pdf', contentId: 'r1', isInline: true, size: 100 },
              { id: 'r2', name: 'r2.png', contentType: 'image/png', isInline: true, size: 100 },
              { id: 'r3', name: 'r3.png', contentType: 'image/png', contentId: '', isInline: true, size: 100 },
              { id: 'r4', name: 'r4.png', contentType: 'image/png', contentId: 'r4', isInline: false, size: 100 },
              { id: 'r5', name: 'r5.png', contentType: 'image/png', contentId: 'r5', size: 100 },
            ],
          }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mPred' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      // No per-attachment bytes were fetched (no third-stage handler in fixture and the test would throw otherwise).
      expect(v.text).not.toContain('data:image/png');
      expect(v.text).not.toContain('data:application/pdf');
      // All five fall into the file-list because none qualify as inline image.
      expect(v.text).toContain('r1.pdf');
      expect(v.text).toContain('r2.png');
      expect(v.text).toContain('r3.png');
      expect(v.text).toContain('r4.png');
      expect(v.text).toContain('r5.png');
    }
  });

  it('convert-mail-to-markdown header lines: Subject-only when To/Cc/Date absent; To+Cc rendered comma-joined when multiple recipients; Cc-only when To absent; receivedDateTime omitted when empty string', async () => {
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');

    // Scenario 1: Subject only, no recipients, empty receivedDateTime string.
    const r1 = await cmd.execute(
      createGraphClient(
        fakeAuth(),
        stagedFetch([
          {
            urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mHA',
            method: 'GET',
            response: () => Response.json({ subject: 'lonely', body: { contentType: 'text', content: 'b' }, receivedDateTime: '', hasAttachments: false }),
          },
        ])
      ),
      { messageId: 'mHA' }
    );
    expect(r1.ok).toBe(true);
    if (r1.ok) {
      const v = r1.value as { text: string };
      expect(v.text).toContain('**Subject:** lonely');
      expect(v.text).not.toContain('**To:**');
      expect(v.text).not.toContain('**Cc:**');
      expect(v.text).not.toContain('**Date:**');
    }

    // Scenario 2: multiple To recipients + Cc (kills the comma-join separator mutant + the recipients-count > 0 boundary).
    const r2 = await cmd.execute(
      createGraphClient(
        fakeAuth(),
        stagedFetch([
          {
            urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mHB',
            method: 'GET',
            response: () =>
              Response.json({
                subject: 'broad',
                toRecipients: [{ emailAddress: { address: 'a@x.com', name: 'Alice' } }, { emailAddress: { address: 'b@x.com' } }, { emailAddress: {} }],
                ccRecipients: [{ emailAddress: { address: 'c@x.com', name: 'Cc1' } }],
                body: { contentType: 'text', content: 'b' },
                hasAttachments: false,
              }),
          },
        ])
      ),
      { messageId: 'mHB' }
    );
    expect(r2.ok).toBe(true);
    if (r2.ok) {
      const v = r2.value as { text: string };
      expect(v.text).toContain('**To:** Alice <a@x.com>, b@x.com');
      expect(v.text).toContain('**Cc:** Cc1 <c@x.com>');
    }

    // Scenario 3: empty toRecipients array yields no `**To:**` line (kills the length > 0 → length >= 0 mutant).
    const r3 = await cmd.execute(
      createGraphClient(
        fakeAuth(),
        stagedFetch([
          {
            urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mHC',
            method: 'GET',
            response: () => Response.json({ subject: 's', toRecipients: [], body: { contentType: 'text', content: 'b' }, hasAttachments: false }),
          },
        ])
      ),
      { messageId: 'mHC' }
    );
    expect(r3.ok).toBe(true);
    if (r3.ok) {
      const v = r3.value as { text: string };
      expect(v.text).not.toContain('**To:**');
    }
  });

  it('convert-mail-to-markdown surfaces the message-level fetch failure unchanged (no body in hand — propagate the GraphError directly to the caller)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mGone',
        method: 'GET',
        response: () => new Response(JSON.stringify({ error: { code: 'ErrorItemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mGone' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(404);
      expect(result.error.message).toContain('gone');
    }
  });

  it('convert-mail-to-markdown header line for **Date:** is rendered when receivedDateTime is a non-empty ISO string', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mDate',
        method: 'GET',
        response: () => Response.json({ subject: 'd', receivedDateTime: '2026-04-30T08:00:00Z', body: { contentType: 'text', content: 'b' }, hasAttachments: false }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mDate' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('**Date:** 2026-04-30T08:00:00Z');
    }
  });

  it('convert-mail-to-markdown formatRecipients returns undefined when toRecipients is non-empty but every entry filters out (no `**To:**` line)', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mEmptyTo',
        method: 'GET',
        response: () =>
          Response.json({ subject: 's', toRecipients: [{ emailAddress: {} }, { emailAddress: {} }], body: { contentType: 'text', content: 'b' }, hasAttachments: false }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mEmptyTo' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).not.toContain('**To:**');
    }
  });

  it('convert-mail-to-markdown skips the per-attachment bytes-fetch when an inline image lacks an id field (Graph metadata anomaly: contentId present but id missing) — cid ref stays in place, no Graph call is made for it', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNoId',
        method: 'GET',
        response: () => Response.json({ subject: 'no-id', body: { contentType: 'html', content: '<p>X: <img src="cid:lost"></p>' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNoId/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [{ name: 'lost.png', contentType: 'image/png', contentId: 'lost', isInline: true, size: 200 }] }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mNoId' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('cid:lost');
      expect(v.text).not.toContain('data:image/png');
    }
  });

  it('convert-mail-to-markdown with hasAttachments:true but an empty attachments-list array — body is rendered without an `**Attachments:**` section, separated from headers by exactly one blank line', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mEmpty',
        method: 'GET',
        response: () => Response.json({ subject: 'empty-list', body: { contentType: 'text', content: 'just-body' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mEmpty/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [] }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mEmpty' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).not.toContain('**Attachments:**');
      expect(v.text).toContain('**Subject:** empty-list\n\njust-body');
    }
  });

  it('convert-mail-to-markdown places the file-attachments list AFTER the body (rendering order is preserved) and prefixes the list with **Attachments:** then closes with the use-`convert-mail-attachment-to-pdf` hint', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mOrder',
        method: 'GET',
        response: () => Response.json({ subject: 'order', body: { contentType: 'text', content: 'body-content-marker' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mOrder/attachments?',
        method: 'GET',
        response: () => Response.json({ value: [{ id: 'a1', name: 'first.pdf', contentType: 'application/pdf', size: 1000, isInline: false }] }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mOrder' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      const bodyIdx = v.text.indexOf('body-content-marker');
      const attachIdx = v.text.indexOf('**Attachments:**');
      expect(bodyIdx).toBeGreaterThanOrEqual(0);
      expect(attachIdx).toBeGreaterThan(bodyIdx);
      expect(v.text).toContain('_Use `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the attachment id to fetch._');
      expect(v.text).toContain('first.pdf');
      expect(v.text).toContain('1.0 KB');
    }
  });

  it('convert-mail-to-markdown meta block declares the expected stable contract (category=mail, GET, producesBytes:true, --message-id is required) — kills mutation on the manifest fields LLM consumers depend on', () => {
    expect(convertMailToMarkdown.meta.category).toBe('mail');
    expect(convertMailToMarkdown.meta.graphMethod).toBe('GET');
    expect(convertMailToMarkdown.meta.producesBytes).toBe(true);
    expect(convertMailToMarkdown.meta.responseShape).toContain('contentType');
    expect(convertMailToMarkdown.meta.responseShape).toContain('text/markdown');
    const messageIdOption = convertMailToMarkdown.meta.options.find((o) => o.name === 'message-id');
    expect(messageIdOption).toBeDefined();
    expect(messageIdOption?.required).toBe(true);
  });

  it('convert-mail-to-markdown attachments-list URL uses the metadata-only $select with the polymorphic-cast for contentId (microsoft.graph.fileAttachment/contentId) — kills mutation on ATTACHMENT_METADATA_SELECT AND prevents the regression where bare `contentId` failed against the base attachment type with `Could not find a property named contentId`', async () => {
    const urls: string[] = [];
    const fetchFn: FetchFn = async (url) => {
      urls.push(url);
      if (url.includes('/attachments?')) {
        return Response.json({ value: [] });
      }
      return Response.json({ subject: 's', body: { contentType: 'text', content: 'b' }, hasAttachments: true });
    };
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    await cmd.execute(graph, { messageId: 'mUrl' });
    const attachmentsUrl = urls.find((u) => u.includes('/attachments?')) ?? '';
    expect(attachmentsUrl).not.toBe('');
    expect(attachmentsUrl).toContain('$select=');
    expect(attachmentsUrl).toContain('id');
    expect(attachmentsUrl).toContain('name');
    expect(attachmentsUrl).toContain('contentType');
    expect(attachmentsUrl).toContain('size');
    expect(attachmentsUrl).toContain('isInline');
    expect(attachmentsUrl).not.toContain('contentBytes');
    // Decode percent-encoding so the bare-`contentId` and cast-form variants
    // can be distinguished by literal substring. The bare form (which Graph
    // rejects on the base attachment type) would appear as `,contentId,` or
    // `,contentId&` after decoding; the cast form carries the
    // `microsoft.graph.fileAttachment/` prefix.
    const decoded = decodeURIComponent(attachmentsUrl);
    expect(decoded).toContain('microsoft.graph.fileAttachment/contentId');
    expect(decoded).not.toMatch(/[,=]contentId([,&]|$)/);
  });

  // Regression: the previous `attachmentMetaSchema.contentId` was
  // `z.string().optional()` — accepts `string | undefined` but REJECTS
  // `null`. Graph's polymorphic-cast response (microsoft.graph.fileAttachment/contentId)
  // returns `contentId: null` on every non-fileAttachment entry, which
  // failed the Zod schema and triggered the `malformed shape` note on every
  // real call. Fixed by relaxing every optional field to `.nullish()`
  // (= `.optional().nullable()`). The downstream `nonEmpty` predicate
  // already handles `null` correctly — narrowed type accepts the slightly
  // wider input.
  it('convert-mail-to-markdown accepts the canonical Graph polymorphic-attachment shape where fileAttachment entries DO carry contentId and other subtypes return contentId: null (regression for the v1.2.0 cast-fix that introduced a Zod-shape false positive)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.includes('/attachments?')) {
        return Response.json({
          value: [
            // fileAttachment — Graph populates contentId.
            {
              '@odata.type': '#microsoft.graph.fileAttachment',
              id: 'a1',
              name: 'report.pdf',
              contentType: 'application/pdf',
              size: 1024,
              isInline: false,
              contentId: null,
            },
            // itemAttachment — Graph returns null on the cast field.
            {
              '@odata.type': '#microsoft.graph.itemAttachment',
              id: 'a2',
              name: 'forwarded.msg',
              contentType: 'application/octet-stream',
              size: 4096,
              isInline: false,
              contentId: null,
            },
          ],
        });
      }
      return Response.json({ subject: 's', body: { contentType: 'text', content: 'b' }, hasAttachments: true });
    };
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mPolymorphic' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { text?: string; note?: string };
    // The markdown body MUST NOT carry the 'malformed shape' note — null
    // contentId is the canonical Graph shape, not a malformed response.
    expect(v.note).toBeUndefined();
    // The Attachments: section should list both files with their ids.
    expect(v.text ?? '').toContain('report.pdf');
    expect(v.text ?? '').toContain('forwarded.msg');
  });

  it("convert-mail-to-markdown --inline-images false skips the per-image bytes fetch entirely (no /attachments/{id} requests) and surfaces inline images in the file-attachments list so the LLM caller doesn't lose visibility (v1.4.0 fresh-pass #6)", async () => {
    const urls: string[] = [];
    const fetchFn: FetchFn = async (url) => {
      urls.push(url);
      if (url.includes('/attachments?')) {
        // Two inline image attachments.
        return Response.json({
          value: [
            { id: 'img1', name: 'logo.png', contentType: 'image/png', size: 1024, isInline: true, contentId: 'logo@cid' },
            { id: 'img2', name: 'banner.jpg', contentType: 'image/jpeg', size: 2048, isInline: true, contentId: 'banner@cid' },
          ],
        });
      }
      // Body references both inline images via cid:.
      return Response.json({
        subject: 'newsletter',
        body: { contentType: 'html', content: '<p>Hello</p><img src="cid:logo@cid"/><img src="cid:banner@cid"/>' },
        hasAttachments: true,
      });
    };
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mInlineSkip', inlineImages: 'false' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { text: string };
    // No per-image bytes fetch happened — `/attachments/{id}` was never hit.
    const perImageFetches = urls.filter((u) => /\/attachments\/img\d/.test(u));
    expect(perImageFetches).toEqual([]);
    // Inline images are now LISTED in the Attachments section (instead of
    // hidden because they were silently embedded).
    expect(v.text).toContain('logo.png');
    expect(v.text).toContain('banner.jpg');
    // Body keeps the raw cid: references (no data: URI substitution).
    expect(v.text).toContain('cid:logo@cid');
    expect(v.text).not.toContain('data:image/png;base64,');
  });

  it('convert-mail-to-markdown default (omitted flag) preserves the historical embedding behaviour — per-image bytes are fetched and data: URIs replace cid: refs', async () => {
    const urls: string[] = [];
    const fetchFn: FetchFn = async (url) => {
      urls.push(url);
      if (url.endsWith('/attachments/img1')) {
        return Response.json({ contentBytes: Buffer.from('fake-png-bytes').toString('base64') });
      }
      if (url.includes('/attachments?')) {
        return Response.json({ value: [{ id: 'img1', name: 'logo.png', contentType: 'image/png', size: 1024, isInline: true, contentId: 'logo@cid' }] });
      }
      return Response.json({ subject: 's', body: { contentType: 'html', content: '<img src="cid:logo@cid"/>' }, hasAttachments: true });
    };
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mInlineDefault' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { text: string };
    // Per-image fetch happened — backward compat preserved.
    expect(urls.some((u) => u.endsWith('/attachments/img1'))).toBe(true);
    expect(v.text).toContain('data:image/png;base64,');
  });

  it('convert-mail-to-markdown rejects --inline-images values other than true/false as a validation_error', async () => {
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fakeFetch({ subject: 's', body: { contentType: 'text', content: 'b' }, hasAttachments: false }));
    const result = await cmd.execute(graph, { messageId: 'mBadFlag', inlineImages: 'yes' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('convert-mail-to-markdown skips the **Subject:** line when the message has no subject field — kills the `if (m.subject !== undefined)` → `if (true)` mutant which would otherwise push `**Subject:** undefined`', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mNoSub',
        method: 'GET',
        response: () => Response.json({ from: { emailAddress: { address: 'x@y.com' } }, body: { contentType: 'text', content: 'just a body' }, hasAttachments: false }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mNoSub' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).not.toContain('**Subject:**');
      expect(v.text).not.toContain('undefined');
      // Output starts with **From:** (the first header that IS present).
      expect(v.text.startsWith('**From:**')).toBe(true);
    }
  });

  it('convert-mail-to-markdown filters inline-image candidates BEFORE the per-attachment bytes fetch — kills the `attachments.filter(isInlineImage)` → `attachments` method-expression mutant by counting that file attachments do NOT trigger an /attachments/{id} fetch', async () => {
    const calls: string[] = [];
    const fetchFn: FetchFn = async (url) => {
      calls.push(url);
      if (url.endsWith('/me/messages/mFilter')) {
        return Response.json({ subject: 'filter', body: { contentType: 'text', content: 'b' }, hasAttachments: true });
      }
      if (url.includes('/attachments?')) {
        return Response.json({
          value: [
            { id: 'aFile', name: 'doc.pdf', contentType: 'application/pdf', isInline: false, size: 1000 },
            { id: 'aOther', name: 'other.docx', contentType: 'application/vnd.openxmlformats', isInline: false, size: 2000 },
          ],
        });
      }
      throw new Error(`unexpected fetch ${url}`);
    };
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mFilter' });
    expect(result.ok).toBe(true);
    // Only 2 calls should have happened: message body + attachments list. NO per-attachment bytes fetch.
    expect(calls.length).toBe(2);
    expect(calls.some((u) => u.includes('/attachments/aFile'))).toBe(false);
    expect(calls.some((u) => u.includes('/attachments/aOther'))).toBe(false);
  });

  it('convert-mail-to-markdown renders the attachments-list with size + contentType + id when present and omits each field when the metadata is missing — and filters out attachments lacking a name', async () => {
    const fetchFn = stagedFetch([
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mList',
        method: 'GET',
        response: () => Response.json({ subject: 'list-shape', body: { contentType: 'text', content: 'body' }, hasAttachments: true }),
      },
      {
        urlPrefix: 'https://graph.microsoft.com/v1.0/me/messages/mList/attachments?',
        method: 'GET',
        response: () =>
          Response.json({
            value: [
              { id: 'aFull', name: 'full.docx', contentType: 'application/vnd.openxmlformats', size: 50_000, isInline: false },
              { id: 'aNoSize', name: 'no-size.pdf', contentType: 'application/pdf', isInline: false },
              { name: 'no-id.txt', contentType: 'text/plain', size: 500, isInline: false },
              { id: 'aNoType', name: 'no-type.bin', size: 999, isInline: false },
              { id: 'aNoName', contentType: 'application/octet-stream', size: 100, isInline: false },
              { id: 'aMB', name: 'exactly-one-mb.bin', contentType: 'application/octet-stream', size: 1_000_000, isInline: false },
            ],
          }),
      },
    ]);
    const cmd = cmdMap['convert-mail-to-markdown'];
    if (!cmd) throw new Error('convert-mail-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'mList' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { text: string };
      expect(v.text).toContain('full.docx');
      expect(v.text).toContain('50.0 KB');
      expect(v.text).toContain('application/vnd.openxmlformats');
      expect(v.text).toContain('id: aFull');
      expect(v.text).toContain('no-size.pdf');
      expect(v.text).toContain('no-id.txt');
      expect(v.text).toContain('500 B');
      expect(v.text).toContain('no-type.bin');
      expect(v.text).not.toContain('aNoName');
      // formatBytes boundary at exactly 1 MB: `< 1_000_000` returns KB; `>= 1_000_000` returns MB.
      // 1_000_000 must render as `1.0 MB` (not `1000.0 KB`) — kills the `n < 1_000_000` → `<=` equality-operator mutant.
      expect(v.text).toContain('exactly-one-mb.bin');
      expect(v.text).toContain('1.0 MB');
      expect(v.text).not.toContain('1000.0 KB');
      // EXACT file-list item shape: name, size in parens with comma-separated contentType, id (kills the `, ` separator string-literal mutants).
      expect(v.text).toContain('- full.docx (50.0 KB, application/vnd.openxmlformats, id: aFull)');
      // Footer hint string verbatim — kills the literal mutant.
      expect(v.text).toContain('_Use `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the attachment id to fetch._');
      // `**Attachments:**` header is present and on its own line preceding the items.
      expect(v.text).toContain('**Attachments:**\n- ');
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
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
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

  it('convert-mail-attachment-to-pdf deletes the empty `.ask-marcel-temp` parent folder after a successful upload-convert-delete cycle (v1.4.0 fresh-pass #7: folder used to linger at OneDrive root)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i7', name: 'plan-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i7') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      // Cleanup: children GET returns empty → folder delete fires.
      if (url.includes('/.ask-marcel-temp:/children')) {
        return Response.json({ value: [] });
      }
      if (url.endsWith('/.ask-marcel-temp') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm7', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    // Both the file delete AND the folder delete happened.
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/items/temp-i7'))).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp'))).toBe(true);
    // The cleanup probed children first.
    expect(calls.some((c) => c.url.includes('/.ask-marcel-temp:/children'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf also sweeps the legacy un-dotted `ask-marcel-temp` folder when empty (QA-008: pre-rename orphan found at a real tenant root)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') return Response.json({ id: 'temp-i8', name: 'plan-temp' });
      if (url.endsWith('/content?format=pdf')) return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      if (url.endsWith('/items/temp-i8') && init?.method === 'DELETE') return new Response(null, { status: 204 });
      // BOTH temp folders report empty → both get the best-effort delete.
      if (url.includes('-temp:/children')) return Response.json({ value: [] });
      if (url.endsWith('-temp') && init?.method === 'DELETE') return new Response(null, { status: 204 });
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm8', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp'))).toBe(true);
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/root:/ask-marcel-temp'))).toBe(true); // the legacy un-dotted orphan
  });

  it('convert-mail-attachment-to-pdf does NOT delete the parent folder when it still contains other files (concurrent invocation safety)', async () => {
    const calls: Array<{ url: string; method?: string }> = [];
    const fetchFn: FetchFn = async (url, init) => {
      calls.push({ url, method: init?.method });
      if (url.endsWith('/attachments/aPdf2')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'b.docx', contentBytes: btoa('docx') });
      }
      if (url.includes(':/content') && init?.method === 'PUT') {
        return Response.json({ id: 'temp-i8', name: 'b-temp' });
      }
      if (url.endsWith('/content?format=pdf')) {
        return new Response(new Uint8Array([0x25, 0x50]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      }
      if (url.endsWith('/items/temp-i8') && init?.method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      // Cleanup: children GET returns one other file → folder delete must NOT fire.
      if (url.includes('/.ask-marcel-temp:/children')) {
        return Response.json({ value: [{ id: 'other-concurrent-upload' }] });
      }
      throw new Error(`unexpected fetch: ${init?.method ?? 'GET'} ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm8', attachmentId: 'aPdf2' });
    expect(result.ok).toBe(true);
    // File delete happened, but the folder delete did NOT — concurrent file survives.
    expect(calls.some((c) => c.method === 'DELETE' && c.url.endsWith('/items/temp-i8'))).toBe(true);
    expect(calls.every((c) => !(c.method === 'DELETE' && c.url.endsWith('/.ask-marcel-temp')))).toBe(true);
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
      const v = result.value as { contentType: string; size: number; base64: string; note: string; passthrough: boolean };
      expect(v.contentType).toBe('text/plain');
      expect(v.note).toContain('pre-checked source');
      expect(v.passthrough).toBe(true); // so --output-path guard blocks writing text into a .pdf (audit A5)
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
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
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

  it('convert-mail-attachment-to-pdf tags a plain-text referenceAttachment short-circuit as passthrough so --output-path cannot write it into a .pdf (audit A5)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefTxt')) {
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/X/notes.txt' });
      }
      if (url.includes('/shares/u!')) {
        return Response.json({ id: 'i-txt', name: 'notes.txt', parentReference: { driveId: 'd1' } });
      }
      if (url.endsWith('/drives/d1/items/i-txt/content')) {
        return new Response('plain text body', { status: 200, headers: { 'content-type': 'text/plain' } });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-pdf'];
    if (!cmd) throw new Error('convert-mail-attachment-to-pdf not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRefTxt' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; passthrough: boolean };
      expect(v.contentType).toBe('text/plain');
      expect(v.passthrough).toBe(true);
    }
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
        return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]), { status: 200, headers: { 'content-type': 'application/pdf' } });
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

  // --- mutation hardening: unconditional error assertions (the `if (error.type === ...)`
  // guards above let type/object/message mutants survive — LESSONS.md), the full
  // IMAGE_EXTENSIONS set, and safeExtension's temp-filename extension via the PUT path. ---
  const pdfCmd = (): { execute: typeof listDrives.execute } => {
    const c = cmdMap['convert-mail-attachment-to-pdf'];
    if (!c) throw new Error('convert-mail-attachment-to-pdf not registered');
    return c;
  };
  const fileAtt =
    (name: string): FetchFn =>
    async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: btoa('bytes') });
      throw new Error(`unexpected ${url}`);
    };
  // Records every PUT (url + decoded body + content-type) and every DELETE url,
  // and serves the convert + cleanup-children responses. `children` controls the
  // cleanup probe body so the `(value ?? [])` default is observable; `putStatus`
  // forces an upload failure.
  type PutRecord = { readonly url: string; readonly body: string; readonly contentType: string };
  const uploadRecorder =
    (name: string, opts: { puts?: PutRecord[]; deletes?: string[]; children?: unknown; putStatus?: number }): FetchFn =>
    async (url, init) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name, contentBytes: btoa('PDFDATA') });
      if (init?.method === 'PUT') {
        const decoded = init.body instanceof Uint8Array ? new TextDecoder().decode(init.body) : String(init.body);
        const ct = (init.headers as Record<string, string> | undefined)?.['content-type'] ?? '';
        opts.puts?.push({ url, body: decoded, contentType: ct });
        if (opts.putStatus !== undefined && opts.putStatus >= 400) return new Response('nope', { status: opts.putStatus });
        return Response.json({ id: 'temp-h1' });
      }
      if (url.endsWith('/content?format=pdf')) return new Response(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { status: 200, headers: { 'content-type': 'application/pdf' } });
      if (url.includes('/.ask-marcel-temp:/children')) return Response.json(opts.children ?? { value: [] });
      if (init?.method === 'DELETE') {
        opts.deletes?.push(url);
        return new Response(null, { status: 204 });
      }
      throw new Error(`unexpected ${init?.method ?? 'GET'} ${url}`);
    };

  it('convert-mail-attachment-to-pdf rejects every IMAGE_EXTENSIONS member with a 415 image hint', async () => {
    for (const ext of ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'svg', 'ico', 'heic', 'heif']) {
      const result = await pdfCmd().execute(createGraphClient(fakeAuth(), fileAtt(`pic.${ext}`)), { messageId: 'm1', attachmentId: 'a1' });
      expect(result.ok).toBe(false);
      if (result.ok) continue;
      expect(result.error.type).toBe('api_error');
      expect(result.error.message).toContain(`${ext} attachment is an image`);
      if (result.error.type === 'api_error') expect(result.error.status).toBe(415);
    }
  });

  it('convert-mail-attachment-to-pdf uploads the decoded bytes under a UUID temp name (safe ext, octet-stream content-type)', async () => {
    const docxPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { puts: docxPuts })), { messageId: 'm1', attachmentId: 'a1' });
    expect(docxPuts[0]?.url).toContain('.ask-marcel-temp/');
    expect(docxPuts[0]?.url.endsWith('.docx') || (docxPuts[0]?.url.includes('.docx:') ?? false)).toBe(true);
    expect(docxPuts[0]?.url).not.toContain('report.docx'); // never the attacker filename
    expect(docxPuts[0]?.body).toBe('PDFDATA'); // decodeBase64 round-trips the payload exactly (kills the loop mutants)
    expect(docxPuts[0]?.contentType).toBe('application/octet-stream'); // kills the L72 content-type literal

    const weirdPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('archive.spreadsheet', { puts: weirdPuts })), { messageId: 'm1', attachmentId: 'a1' });
    expect(weirdPuts[0]?.url.endsWith('.bin') || (weirdPuts[0]?.url.includes('.bin:') ?? false)).toBe(true); // ext > 8 chars → bin

    const noExtPuts: PutRecord[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('README', { puts: noExtPuts })), { messageId: 'm1', attachmentId: 'a1' });
    // 'README' is ≤8 alnum chars: a broken no-dot guard would slice the whole
    // name → '.readme'. The correct guard returns 'bin'.
    expect(noExtPuts[0]?.url.endsWith('.bin') || (noExtPuts[0]?.url.includes('.bin:') ?? false)).toBe(true);
  });

  it('convert-mail-attachment-to-pdf returns an empty raw-bytes envelope for a plain-text attachment with no contentBytes', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt' }); // no contentBytes field
    const result = await pdfCmd().execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const env = result.value as { contentType: string; size: number; base64: string; note: string };
    expect(env.contentType).toBe('text/plain');
    expect(env.size).toBe(0); // empty contentBytes default '' → 0 bytes (mutant default is non-empty)
    expect(env.base64).toBe('');
    expect(env.note).toContain('note.txt');
  });

  it('convert-mail-attachment-to-pdf deletes the temp file AND the temp folder when the cleanup probe shows it empty', async () => {
    const deletes: string[] = [];
    await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { deletes, children: {} })), { messageId: 'm1', attachmentId: 'a1' });
    expect(deletes.some((u) => u.includes('/me/drive/items/temp-h1'))).toBe(true); // temp file always deleted
    // children body has no `value` → the `?? []` default makes it empty → folder
    // deleted. A non-empty default mutant would skip this delete.
    expect(deletes.some((u) => u.endsWith('/.ask-marcel-temp'))).toBe(true);
  });

  it('convert-mail-attachment-to-pdf propagates an upload (PUT) failure rather than mislabelling it as a missing id', async () => {
    const puts: PutRecord[] = [];
    const result = await pdfCmd().execute(createGraphClient(fakeAuth(), uploadRecorder('report.docx', { puts, putStatus: 403 })), { messageId: 'm1', attachmentId: 'a1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(403); // mutant skips the guard → 500 'no driveItem id'
    expect(result.error.message).not.toContain('upload returned no driveItem id');
  });

  it('convert-mail-attachment-to-pdf propagates the initial attachment-fetch failure', async () => {
    const result = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => new Response('nope', { status: 404 })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(404); // mutant skips the guard → reads undefined.value → throws
  });

  it('convert-mail-attachment-to-pdf errs unconditionally for itemAttachment, missing @odata.type, and unknown type', async () => {
    const item = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: {} })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(item.ok).toBe(false);
    if (!item.ok) {
      expect(item.error.type).toBe('api_error');
      expect(item.error.message).toContain('convert-mail-attachment-to-markdown');
      if (item.error.type === 'api_error') expect(item.error.status).toBe(400);
    }
    const noType = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ name: 'x.docx' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(noType.ok).toBe(false);
    if (!noType.ok) {
      expect(noType.error.type).toBe('api_error');
      expect(noType.error.message).toContain('missing @odata.type discriminator');
      if (noType.error.type === 'api_error') expect(noType.error.status).toBe(400);
    }
    const unknown = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.weird' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) {
      expect(unknown.error.type).toBe('api_error');
      expect(unknown.error.message).toContain('unsupported attachment type: #microsoft.graph.weird');
      if (unknown.error.type === 'api_error') expect(unknown.error.status).toBe(400);
    }
  });

  it('convert-mail-attachment-to-pdf errs unconditionally on referenceAttachment edge cases (missing/empty sourceUrl, resolved missing id, resolved missing driveId, image source) and on resolve / upload failures', async () => {
    // missing sourceUrl (undefined)
    const noUrl = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(noUrl.ok).toBe(false);
    if (!noUrl.ok) {
      expect(noUrl.error.type).toBe('api_error');
      expect(noUrl.error.message).toContain('referenceAttachment missing sourceUrl');
      if (noUrl.error.type === 'api_error') expect(noUrl.error.status).toBe(400);
    }

    // empty-string sourceUrl exercises the `=== ''` operand specifically
    const emptyUrl = await pdfCmd().execute(
      createGraphClient(fakeAuth(), async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: '' })),
      { messageId: 'm1', attachmentId: 'a1' }
    );
    expect(emptyUrl.ok).toBe(false);
    if (!emptyUrl.ok) {
      expect(emptyUrl.error.type).toBe('api_error');
      expect(emptyUrl.error.message).toContain('referenceAttachment missing sourceUrl');
    }

    // resolved driveItem with no id (driveId present → the itemId operand is the deciding one)
    const itemIdMissing: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return Response.json({ name: 'q.docx', parentReference: { driveId: 'd1' } });
    };
    const iim = await pdfCmd().execute(createGraphClient(fakeAuth(), itemIdMissing), { messageId: 'm1', attachmentId: 'a1' });
    expect(iim.ok).toBe(false);
    if (!iim.ok) {
      expect(iim.error.type).toBe('api_error');
      expect(iim.error.message).toContain('resolved driveItem missing id or driveId');
    }

    // resolved driveItem with no driveId (the first operand)
    const driveIdMissing: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return Response.json({ id: 'i1', name: 'q.docx' });
    };
    const dim = await pdfCmd().execute(createGraphClient(fakeAuth(), driveIdMissing), { messageId: 'm1', attachmentId: 'a1' });
    expect(dim.ok).toBe(false);
    if (!dim.ok) {
      expect(dim.error.type).toBe('api_error');
      expect(dim.error.message).toContain('resolved driveItem missing id or driveId');
      if (dim.error.type === 'api_error') expect(dim.error.status).toBe(500);
    }

    // resolved driveItem that is an image → 415 (reference-side image guard)
    const refImage: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/photo.png' });
      return Response.json({ id: 'i1', name: 'photo.png', parentReference: { driveId: 'd1' } });
    };
    const ri = await pdfCmd().execute(createGraphClient(fakeAuth(), refImage), { messageId: 'm1', attachmentId: 'a1' });
    expect(ri.ok).toBe(false);
    if (!ri.ok) {
      expect(ri.error.type).toBe('api_error');
      expect(ri.error.message).toContain('png attachment is an image');
      if (ri.error.type === 'api_error') expect(ri.error.status).toBe(415);
    }

    // /shares resolve failure propagates
    const resolveFail: FetchFn = async (url) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://x/q.docx' });
      return new Response('nope', { status: 404 });
    };
    const rf = await pdfCmd().execute(createGraphClient(fakeAuth(), resolveFail), { messageId: 'm1', attachmentId: 'a1' });
    expect(rf.ok).toBe(false);
    if (rf.ok) return;
    expect(rf.error.type).toBe('api_error');
    if (rf.error.type === 'api_error') expect(rf.error.status).toBe(404);
  });

  it('convert-mail-attachment-to-pdf errs unconditionally when the upload returns no driveItem id', async () => {
    const noUploadId: FetchFn = async (url, init) => {
      if (url.includes('/attachments/')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'r.docx', contentBytes: btoa('hi') });
      if (init?.method === 'PUT') return Response.json({ name: 'no-id-returned' }); // upload response without `id`
      throw new Error(`unexpected ${url}`);
    };
    const uploadNoId = await pdfCmd().execute(createGraphClient(fakeAuth(), noUploadId), { messageId: 'm1', attachmentId: 'a1' });
    expect(uploadNoId.ok).toBe(false);
    if (!uploadNoId.ok) {
      expect(uploadNoId.error.type).toBe('api_error');
      expect(uploadNoId.error.message).toContain('upload returned no driveItem id');
      if (uploadNoId.error.type === 'api_error') expect(uploadNoId.error.status).toBe(500);
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

  it('convert-mail-attachment-to-markdown content-sniffs a UTF-8 fileAttachment to an inline text envelope', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README.md', contentBytes: btoa('# Hi') });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aText' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text?: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toBe('# Hi');
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

  it('convert-mail-attachment-to-markdown extracts `## OpenDocument metadata` for an odt fileAttachment when --include-metadata true is set', async () => {
    const odtBytes = await buildRichOdt();
    let binary = '';
    for (const byte of odtBytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.odt', contentBytes: btoa(binary) });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aOdt', includeMetadata: 'true' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('## OpenDocument metadata');
      expect(v.text).toContain('Q4 Plan');
    }
  });

  it('convert-mail-attachment-to-markdown converts an odt fileAttachment body without --include-metadata (no metadata block)', async () => {
    const odtBytes = await buildRichOdt();
    let binary = '';
    for (const byte of odtBytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'plan.odt', contentBytes: btoa(binary) });
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aOdtNoMeta' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { contentType: string; text: string };
    expect(v.text).toContain('# Heading One');
    expect(v.text).not.toContain('## OpenDocument metadata');
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

  it('extract-mail-attachment-images returns a media envelope for a pptx fileAttachment', async () => {
    const bytes = await buildMediaSamples();
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aDeck')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentBytes: btoa(binary) });
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aDeck' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { count: number; media: ReadonlyArray<{ path: string }> };
      expect(v.count).toBe(4);
      expect(v.media.map((m) => m.path)).toEqual(['ppt/media/diagram.gif', 'word/media/chart.svg', 'word/media/image1.png', 'xl/media/photo.jpeg']);
    }
  });

  it('extract-mail-attachment-images resolves a referenceAttachment via /shares and extracts the linked docx images', async () => {
    const bytes = await buildMediaSamples();
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRef'))
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x/q3.docx' });
      if (url.includes('/shares/u!')) return Response.json({ id: 'i9', name: 'q3.docx', parentReference: { driveId: 'd9' } });
      if (url.includes('/drives/d9/items/i9/content')) return new Response(bytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/octet-stream' } });
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRef' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { count: number; media: ReadonlyArray<{ path: string }> };
      expect(v.count).toBe(4);
      expect(v.media.map((m) => m.path)).toContain('word/media/image1.png');
    }
  });

  it('extract-mail-attachment-images returns a media envelope for an xlsx fileAttachment (xlsx family)', async () => {
    const bytes = await buildMediaSamples();
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'sheet.xlsx', contentBytes: btoa(binary) });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aXlsx' });
    expect(result.ok).toBe(true);
    if (result.ok) expect((result.value as { count: number }).count).toBe(4);
  });

  it('extract-mail-attachment-images extracts PNG page images from a PDF fileAttachment', async () => {
    const contentBytes = Buffer.from(buildPdfWithImage()).toString('base64');
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'scan.pdf', contentBytes });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { count: number; media: ReadonlyArray<{ path: string }> };
    expect(v.count).toBe(1);
    expect(v.media[0]?.path).toMatch(/^pdf\/page1\/.+\.png$/);
  });

  it('extract-mail-attachment-images treats a trailing-dot attachment name as non-OOXML', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'weird.', contentBytes: btoa('zzz') });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aDot' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') expect(result.error.message).toContain('<no-extension> is not a supported document');
  });

  it('extract-mail-attachment-images surfaces a media-extraction failure for an OOXML-named attachment that is not a valid zip', async () => {
    let binary = '';
    for (const byte of buildMalformedDocx()) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'corrupt.docx', contentBytes: btoa(binary) });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aBad' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('ooxml media extraction failed');
  });

  it('extract-mail-attachment-images propagates the attachment-GET error (404) before inspecting @odata.type', async () => {
    const fetchFn: FetchFn = async () => new Response('gone', { status: 404 });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aGone' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(404);
  });

  it('extract-mail-attachment-images returns a 415 api_error for an itemAttachment (no embedded OOXML doc)', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.itemAttachment', item: { '@odata.type': '#microsoft.graph.contact' } });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aItem' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') {
      expect(result.error.status).toBe(415);
      expect(result.error.message).toContain('itemAttachment');
    }
  });

  it('extract-mail-attachment-images returns api_error for an unknown attachment @odata.type', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.weirdNewType' });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aWeird' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('unsupported attachment type');
  });

  it('extract-mail-attachment-images returns api_error when the outer attachment is missing @odata.type', async () => {
    const fetchFn: FetchFn = async () => Response.json({ name: 'x.docx' });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aNoType' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('missing @odata.type discriminator');
  });

  it('extract-mail-attachment-images rejects a referenceAttachment with no sourceUrl', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment' });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRefBad' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') {
      expect(result.error.status).toBe(400);
      expect(result.error.message).toContain('missing sourceUrl');
    }
  });

  it('extract-mail-attachment-images rejects a referenceAttachment with an empty-string sourceUrl', async () => {
    const fetchFn: FetchFn = async () => Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: '' });
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aEmptyUrl' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('missing sourceUrl');
  });

  it('extract-mail-attachment-images propagates a failed /shares resolution', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRefErr'))
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x/q3.docx' });
      if (url.includes('/shares/u!')) return new Response('forbidden', { status: 403 });
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRefErr' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.status).toBe(403);
  });

  it('extract-mail-attachment-images errs when a resolved referenceAttachment lacks driveId/id', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aRef2'))
        return Response.json({ '@odata.type': '#microsoft.graph.referenceAttachment', sourceUrl: 'https://contoso.sharepoint.com/sites/x/q3.docx' });
      if (url.includes('/shares/u!')) return Response.json({ id: 'i9' });
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', attachmentId: 'aRef2' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('missing id or driveId');
  });

  it('extract-mail-attachment-images returns a validation_error when messageId is missing', async () => {
    const cmd = cmdMap['extract-mail-attachment-images'];
    if (!cmd) throw new Error('extract-mail-attachment-images not registered');
    const result = await cmd.execute(
      createGraphClient(fakeAuth(), async () => Response.json({})),
      { attachmentId: 'a1' }
    );
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
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

  it('convert-mail-attachment-to-markdown extracts pptx slide text from a pptx fileAttachment', async () => {
    let bin = '';
    for (const b of await buildRichPptx()) bin += String.fromCharCode(b);
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPptx')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentBytes: btoa(bin) });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPptx' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('## Slide 1');
      expect(v.text).toContain('Quarterly Review');
    }
  });

  it('convert-mail-attachment-to-markdown aliases a macro-enabled .docm fileAttachment onto the docx path', async () => {
    const docxBytes = await buildSampleDocx();
    let binary = '';
    for (const byte of docxBytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aDocm')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'macro.docm', contentBytes: btoa(binary) });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aDocm' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('# Sample Heading');
    }
  });

  it('convert-mail-attachment-to-markdown extracts a `## PPTX metadata` document for a pptx fileAttachment when --include-metadata true is set', async () => {
    const pptxBytes = await buildRichPptx();
    let binary = '';
    for (const byte of pptxBytes) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPptxMeta')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pptx', contentBytes: btoa(binary) });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPptxMeta', includeMetadata: 'true' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/markdown');
      expect(v.text).toContain('## PPTX metadata');
      expect(v.text).toContain('Quarterly Review');
    }
  });

  it('convert-mail-attachment-to-markdown extracts a born-digital PDF fileAttachment’s text layer (text/plain)', async () => {
    let bin = '';
    for (const b of buildPdfWithText()) bin += String.fromCharCode(b);
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aPdf')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: btoa(bin) });
      }
      throw new Error(`unexpected ${url}`);
    };
    const cmd = cmdMap['convert-mail-attachment-to-markdown'];
    if (!cmd) throw new Error('convert-mail-attachment-to-markdown not registered');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { messageId: 'm1', attachmentId: 'aPdf' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { contentType: string; text: string };
      expect(v.contentType).toBe('text/plain');
      expect(v.text).toContain('Hello from the');
    }
  });

  it('convert-mail-attachment-to-markdown errs with `<no-extension>` placeholder for a dotless-name BINARY fileAttachment (content-sniff fails)', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/attachments/aNoExt')) {
        return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'README', contentBytes: btoa(String.fromCharCode(0xff, 0xfe, 0xfd)) });
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

describe('calendar event attachments', () => {
  it('list-calendar-event-attachments hits /me/events/{id}/attachments with a slim default --select, and honors an explicit --select', async () => {
    let defaultUrl = '';
    await cmdMap['list-calendar-event-attachments']?.execute(
      createGraphClient(fakeAuth(), async (url) => {
        defaultUrl = url;
        return Response.json({ value: [] });
      }),
      { eventId: 'e1' }
    );
    expect(defaultUrl).toContain('/me/events/e1/attachments');
    expect(defaultUrl).toContain('isInline');
    let overrideUrl = '';
    await cmdMap['list-calendar-event-attachments']?.execute(
      createGraphClient(fakeAuth(), async (url) => {
        overrideUrl = url;
        return Response.json({ value: [] });
      }),
      { eventId: 'e1', select: 'id' }
    );
    expect(overrideUrl).not.toContain('isInline');
  });

  it('list-calendar-event-attachments returns a validation_error when eventId is missing', async () => {
    const result = await cmdMap['list-calendar-event-attachments']?.execute(createGraphClient(fakeAuth(), fakeFetch({ value: [] })), {});
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('convert-calendar-event-attachment-to-markdown converts a docx event attachment and threads includeMetadata (true appends the metadata block; false/absent omits it)', async () => {
    const docx = await buildSampleDocx();
    let binary = '';
    for (const byte of docx) binary += String.fromCharCode(byte);
    const fetchFn: FetchFn = async (url) => {
      if (url.includes('/me/events/e1/attachments/a1')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'r.docx', contentBytes: btoa(binary) });
      throw new Error(`unexpected ${url}`);
    };
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const text = async (params: Record<string, string>): Promise<string> => {
      const r = await cmdMap['convert-calendar-event-attachment-to-markdown']?.execute(graph, params);
      return r?.ok ? (r.value as { text: string }).text : '';
    };
    expect(await text({ eventId: 'e1', attachmentId: 'a1' })).toContain('# Sample Heading');
    expect(await text({ eventId: 'e1', attachmentId: 'a1' })).not.toContain('## DOCX metadata');
    expect(await text({ eventId: 'e1', attachmentId: 'a1', includeMetadata: 'true' })).toContain('## DOCX metadata');
    // `false` must be an accepted enum value (kills the enum-literal mutant) and must omit metadata
    const explicitFalse = await cmdMap['convert-calendar-event-attachment-to-markdown']?.execute(graph, { eventId: 'e1', attachmentId: 'a1', includeMetadata: 'false' });
    expect(explicitFalse?.ok).toBe(true);
    if (explicitFalse?.ok) expect((explicitFalse.value as { text: string }).text).not.toContain('## DOCX metadata');
  });

  it('convert-calendar-event-attachment-to-markdown returns a validation_error when attachmentId is missing', async () => {
    const result = await cmdMap['convert-calendar-event-attachment-to-markdown']?.execute(createGraphClient(fakeAuth(), fakeFetch({})), { eventId: 'e1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('convert-calendar-event-attachment-to-pdf short-circuits a pdf event attachment to raw bytes through the shared pipeline', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.includes('/me/events/e1/attachments/a1')) return Response.json({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'deck.pdf', contentBytes: btoa('PDFB') });
      throw new Error(`unexpected ${url}`);
    };
    const result = await cmdMap['convert-calendar-event-attachment-to-pdf']?.execute(createGraphClient(fakeAuth(), fetchFn), { eventId: 'e1', attachmentId: 'a1' });
    expect(result?.ok).toBe(true);
    if (result?.ok) {
      const v = result.value as { contentType: string; note?: string };
      expect(v.contentType).toBe('application/pdf');
      expect(String(v.note)).toContain('raw bytes');
    }
  });

  it('convert-calendar-event-attachment-to-pdf returns a validation_error when eventId is missing', async () => {
    const result = await cmdMap['convert-calendar-event-attachment-to-pdf']?.execute(createGraphClient(fakeAuth(), fakeFetch({})), { attachmentId: 'a1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });
});

describe('convert-drive-item-zip', () => {
  const zipResponse =
    (bytes: Uint8Array): FetchFn =>
    async () =>
      new Response(bytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/zip' } });

  it('converts every Office/text entry, notes a malformed entry, and skips non-convertible ones', async () => {
    const archive = await buildSampleZipArchive();
    const result = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), zipResponse(archive)), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { count: number; truncated?: boolean; files: ReadonlyArray<{ path: string; contentType?: string; text?: string; note?: string }> };
    const at = (p: string): { contentType?: string; text?: string; note?: string } => v.files.find((f) => f.path === p) ?? {};
    expect(v.count).toBe(17);
    expect(v.truncated).toBeUndefined(); // under the cap → no truncation flag
    expect(at('report.docx').text).toContain('# Sample Heading');
    expect(at('report.docx').text).not.toContain('## DOCX metadata'); // no metadata block without the flag
    expect(at('data.xlsx').text).toContain('## Sheet1');
    expect(at('deck.pptx').text).toContain('## Slide 1'); // pptx slide text extracted inline (no metadata block without the flag)
    expect(at('deck.pptx').text).not.toContain('## PPTX metadata');
    expect(at('plan.odt').text).toContain('# Heading One');
    expect(at('notes.txt').text).toBe('hello from the archive');
    expect(at('notes.txt').contentType).toBe('text/plain');
    expect(at('broken.docx').note).toContain('conversion failed');
    // A born-digital pdf entry → its text layer is extracted inline.
    expect(at('scan.pdf').text).toContain('Hello from the');
    expect(at('scan.pdf').contentType).toBe('text/plain');
    // A pdf with no text layer (scanned / image-only) → noted, not failed.
    expect(at('blank.pdf').note).toContain('no extractable text layer');
    // Legacy OLE formats: .xls → sheetjs markdown table, .doc → text/plain body, .ppt → convert-to-pdf note.
    expect(at('legacy.xls').text).toContain('## Legacy');
    expect(at('legacy.doc').text).toContain('Hello from the legacy doc');
    expect(at('legacy.doc').contentType).toBe('text/plain');
    expect(at('corrupt.doc').note).toContain('text extraction failed'); // word-extractor throws on non-OLE bytes → noted, not failed
    expect(at('legacy.ppt').note).toContain('convert it to PDF');
    // An Outlook .msg entry → rendered to markdown (headers + body) with its own
    // attachment ("summary.txt") recursed through the same dispatch and inlined.
    expect(at('mail.msg').text).toContain('# Quarterly Report');
    expect(at('mail.msg').text).toContain('**From:** Jordan Avery <jordan.avery@example.com>');
    expect(at('mail.msg').text).toContain('### summary.txt');
    expect(at('mail.msg').text).toContain('Attachment body text');
    expect(at('mail.msg').contentType).toBe('text/markdown');
    expect(at('photo.png').note).toContain('png is an image'); // raster image → noted, not unpacked
    // QA-007: entries INSIDE a container get container-neutral guidance — the
    // sibling commands can only address top-level drive items, never zip entries.
    expect(at('photo.png').note).toContain('extract it from the archive/message');
    expect(at('photo.png').note).not.toContain('extract-drive-item-images');
    expect(at('blank.pdf').note).not.toContain('download-drive-item-as-pdf');
    expect(at('data.bin').note).toContain('bin is not a convertible'); // binary bytes → skip note
    // A dotless entry with valid-UTF-8 bytes now content-sniffs to text (no extension list to consult).
    expect(at('LICENSE').text).toBe('a dotless, no-extension entry');
    expect(at('LICENSE').contentType).toBe('text/plain');
    // A dotless entry with BINARY bytes → skipped via the `ext === ''` → "<no-extension>" branch.
    expect(at('rawblob').note).toContain('<no-extension>');
    expect(at('rawblob').text).toBeUndefined();
  });

  it('appends each Office file’s metadata block when --include-metadata true (and accepts an explicit false)', async () => {
    const archive = await buildSampleZipArchive();
    const withMeta = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), zipResponse(archive)), { driveId: 'd1', itemId: 'i1', includeMetadata: 'true' });
    expect(withMeta?.ok).toBe(true);
    if (!withMeta?.ok) return;
    const files = (withMeta.value as { files: ReadonlyArray<{ path: string; text?: string }> }).files;
    expect(files.find((f) => f.path === 'report.docx')?.text).toContain('## DOCX metadata');
    expect(files.find((f) => f.path === 'data.xlsx')?.text).toContain('## Workbook metadata');
    expect(files.find((f) => f.path === 'deck.pptx')?.text).toContain('## PPTX metadata');
    expect(files.find((f) => f.path === 'plan.odt')?.text).toContain('## OpenDocument metadata');
    // `false` is an accepted enum value and must omit the metadata blocks
    const explicitFalse = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), zipResponse(await buildSampleZipArchive())), {
      driveId: 'd1',
      itemId: 'i1',
      includeMetadata: 'false',
    });
    expect(explicitFalse?.ok).toBe(true);
    if (explicitFalse?.ok)
      expect((explicitFalse.value as { files: ReadonlyArray<{ path: string; text?: string }> }).files.find((f) => f.path === 'report.docx')?.text).not.toContain(
        '## DOCX metadata'
      );
  });

  it('caps at 100 entries and flags truncation', async () => {
    const archive = await buildOversizedZipArchive();
    const result = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), zipResponse(archive)), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { count: number; truncated?: boolean; totalEntries?: number };
    expect(v.count).toBe(100);
    expect(v.truncated).toBe(true);
    expect(v.totalEntries).toBe(101);
  });

  it('propagates a zip-parse error for non-zip bytes', async () => {
    const result = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), zipResponse(new Uint8Array([1, 2, 3]))), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.error.type).toBe('api_error');
      expect(result.error.message).toContain('not a valid zip archive');
    }
  });

  it('propagates the content-fetch error', async () => {
    const fetchFn: FetchFn = async () =>
      new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } });
    const result = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.error.type).toBe('api_error');
      // the fetch error is returned as-is — NOT swallowed and re-surfaced as a zip-parse failure
      expect(result.error.message).not.toContain('not a valid zip archive');
    }
  });

  it('returns a validation_error when itemId is missing', async () => {
    const result = await cmdMap['convert-drive-item-zip']?.execute(createGraphClient(fakeAuth(), fakeFetch({})), { driveId: 'd1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });
});

describe('extract-sharepoint-links-in-documents', () => {
  // Serve the document bytes for the `/content` fetch and a resolved driveItem
  // (or a caller-supplied response) for each `/shares/{token}/driveItem` resolve.
  const docResponse =
    (bytes: Uint8Array, shareResolver?: () => Response): FetchFn =>
    async (url) => {
      if (url.includes('/items/i1/content'))
        return new Response(bytes as unknown as BodyInit, { status: 200, headers: { 'content-type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' } });
      if (url.includes('/shares/'))
        return shareResolver
          ? shareResolver()
          : Response.json({ id: 'item-1', name: 'spec.docx', webUrl: 'https://contoso.sharepoint.com/sites/Marketing/spec', parentReference: { driveId: 'drive-X' } });
      throw new Error(`unexpected ${url}`);
    };

  it('extracts and resolves SharePoint links from a docx, dedupes, and filters non-SharePoint externals', async () => {
    const bytes = await buildDocxWithSharepointLinks();
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), docResponse(bytes)), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as {
      driveId: string;
      itemId: string;
      links: ReadonlyArray<{ url: string; driveId?: string; itemId?: string; name?: string; webUrl?: string; error?: string }>;
      truncated: boolean;
      skippedCount: number;
    };
    expect(v.driveId).toBe('d1');
    expect(v.itemId).toBe('i1');
    // two identical SharePoint rels deduped to one; the example.com external is filtered out; the internal styles rel never surfaces
    expect(v.links).toHaveLength(1);
    expect(v.links[0]?.url).toBe('https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/spec.docx');
    expect(v.links[0]?.driveId).toBe('drive-X');
    expect(v.links[0]?.itemId).toBe('item-1');
    expect(v.links[0]?.name).toBe('spec.docx');
    expect(v.links.some((l) => l.url.includes('example.com'))).toBe(false);
    expect(v.truncated).toBe(false);
    expect(v.skippedCount).toBe(0);
  });

  it('extracts and resolves SharePoint links from an OpenDocument (.odt) via content.xml xlink:href, dedupes, and filters non-SharePoint', async () => {
    const bytes = await buildOdtWithSharepointLinks();
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), docResponse(bytes)), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { links: ReadonlyArray<{ url: string; driveId?: string }> };
    expect(v.links).toHaveLength(1); // the SharePoint href appears twice (deduped); example.com filtered out
    expect(v.links[0]?.url).toBe('https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/odf-spec.odt');
    expect(v.links[0]?.driveId).toBe('drive-X');
    expect(v.links.some((l) => l.url.includes('example.com'))).toBe(false);
  });

  it('returns an empty links array for a valid Office doc with no external SharePoint links', async () => {
    const bytes = await buildSampleDocx();
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), docResponse(bytes)), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(true);
    if (result?.ok) expect((result.value as { links: ReadonlyArray<unknown> }).links).toHaveLength(0);
  });

  it('returns an api_error 400 when the item is not a zip-based Office document', async () => {
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), docResponse(new Uint8Array([0x25, 0x50, 0x44, 0x46]))), {
      driveId: 'd1',
      itemId: 'i1',
    });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.error.type).toBe('api_error');
      expect(result.error.message).toContain('not a zip-based Office document');
      if (result.error.type === 'api_error') expect(result.error.status).toBe(400);
    }
  });

  it('captures a per-link resolve failure inside the entry instead of failing the whole call', async () => {
    const bytes = await buildDocxWithSharepointLinks();
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(
      createGraphClient(
        fakeAuth(),
        docResponse(
          bytes,
          () => new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } })
        )
      ),
      { driveId: 'd1', itemId: 'i1' }
    );
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { links: ReadonlyArray<{ url: string; error?: string }> };
    expect(v.links).toHaveLength(1);
    expect(v.links[0]?.error).toBeDefined();
  });

  it('propagates the content-fetch error rather than mislabelling it as a non-OOXML document', async () => {
    const fetchFn: FetchFn = async () =>
      new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } });
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), fetchFn), { driveId: 'd1', itemId: 'i1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.error.type).toBe('api_error');
      // the original fetch error is returned as-is — NOT swallowed and re-surfaced as the non-zip 400
      expect(result.error.message).not.toContain('not a zip-based Office document');
    }
  });

  it('returns a validation_error when itemId is missing', async () => {
    const result = await cmdMap['extract-sharepoint-links-in-documents']?.execute(createGraphClient(fakeAuth(), fakeFetch({})), { driveId: 'd1' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('validation_error');
  });
});

describe('sharepoint-link-extractor helpers', () => {
  it('buildShareToken encodes a url as u! + unpadded base64url (+ → -, / → _, = stripped)', () => {
    // url chosen so its raw base64 contains +, /, AND = padding — exercises every replaceAll
    expect(buildShareToken('https://contoso.sharepoint.com/x?q=~//?>')).toBe('u!aHR0cHM6Ly9jb250b3NvLnNoYXJlcG9pbnQuY29tL3g_cT1-Ly8_Pg');
  });

  const okShares: FetchFn = async (url) => {
    if (url.includes('/shares/')) return Response.json({ id: 'i', name: 'n', parentReference: { driveId: 'd' } });
    throw new Error(`unexpected ${url}`);
  };
  const urls = (n: number): string[] => Array.from({ length: n }, (_, i) => `https://contoso.sharepoint.com/doc${i}`);

  it('resolveSharepointUrls keeps exactly MAX_LINKS (25) without truncating, and truncates at 26', async () => {
    const graph = createGraphClient(fakeAuth(), okShares);
    const at25 = await resolveSharepointUrls(graph, urls(25));
    expect(at25.links).toHaveLength(25);
    expect(at25.truncated).toBe(false); // boundary: 25 > 25 is false
    expect(at25.skippedCount).toBe(0);

    const at26 = await resolveSharepointUrls(graph, urls(26));
    expect(at26.links).toHaveLength(25);
    expect(at26.truncated).toBe(true);
    expect(at26.skippedCount).toBe(1);
  });

  it('resolveSharepointUrls returns the bare message for an api_error and a type-labelled message for a non-api_error', async () => {
    const apiFail = createGraphClient(
      fakeAuth(),
      async () => new Response(JSON.stringify({ error: { code: 'accessDenied', message: 'no access here' } }), { status: 403, headers: { 'content-type': 'application/json' } })
    );
    const apiRes = await resolveSharepointUrls(apiFail, ['https://contoso.sharepoint.com/x']);
    expect(apiRes.links[0]?.error).toBeDefined();
    expect(apiRes.links[0]?.error?.startsWith('api_error:')).toBe(false); // api_error → bare message, no type prefix

    const netFail = createGraphClient(fakeAuth(), async () => {
      throw new Error('socket hang up');
    });
    const netRes = await resolveSharepointUrls(netFail, ['https://contoso.sharepoint.com/y']);
    expect(netRes.links[0]?.error).toContain('network_error:'); // non-api_error → labelled `${type}: ${message}`
  });

  it('resolveSharepointUrls tolerates a resolved driveItem with no parentReference (driveId undefined, no throw)', async () => {
    const noParent = createGraphClient(fakeAuth(), async (url) => (url.includes('/shares/') ? Response.json({ id: 'i1', name: 'n1' }) : new Response('x', { status: 500 })));
    const res = await resolveSharepointUrls(noParent, ['https://contoso.sharepoint.com/z']);
    expect(res.links).toHaveLength(1);
    expect(res.links[0]?.itemId).toBe('i1');
    expect(res.links[0]?.driveId).toBeUndefined();
    expect(res.links[0]?.error).toBeUndefined();
  });
});

describe('mail-quote-stripper', () => {
  // Each fixture puts a multi-char attribute (`data-x="y"`) before the id/class
  // so the `<tag[^>]*` quantifier must match more than one char (kills the
  // `[^>]*` → `[^>]` regex mutants), and the Gmail ones prefix the class value
  // with `im ` so `class="[^"]*gmail_quote` must match a non-empty prefix
  // (kills the `[^"]*` → `["]*` mutants).
  const boundaryFixtures = [
    { label: 'Outlook divRplyFwdMsg', html: '<p>reply body</p><div data-x="y" id="divRplyFwdMsg"><b>From:</b> Alice</div><p>old quoted</p>' },
    { label: 'Outlook appendonsend', html: '<p>reply body</p><div data-x="y" id="appendonsend"></div><blockquote>old quoted</blockquote>' },
    { label: 'Outlook mobile container', html: '<p>reply body</p><div data-x="y" id="mail-editor-reference-message-container"><p>old quoted</p></div>' },
    { label: 'Outlook stopSpelling hr', html: '<p>reply body</p><hr data-x="y" id="stopSpelling"><p>old quoted</p>' },
    { label: 'Gmail div quote', html: '<p>reply body</p><div data-x="y" class="im gmail_quote">old quoted</div>' },
    { label: 'Gmail blockquote quote', html: '<p>reply body</p><blockquote data-x="y" class="im gmail_quote">old quoted</blockquote>' },
  ];
  it.each(boundaryFixtures)('stripQuotedReplies cuts the body at the $label boundary (with attributes before the marker)', ({ html }) => {
    const r = stripQuotedReplies(html);
    expect(r.stripped).toBe(true);
    expect(r.html).toBe(`<p>reply body</p>${'<p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>'}`);
    expect(r.html).not.toContain('old quoted');
  });

  it('stripQuotedReplies leaves a body with no quote markers untouched (a bare blockquote is NOT a marker)', () => {
    const html = '<p>Just a normal email with a <blockquote>short inline citation</blockquote>.</p>';
    const r = stripQuotedReplies(html);
    expect(r.stripped).toBe(false);
    expect(r.html).toBe(html);
  });

  it('stripQuotedReplies cuts at the EARLIEST marker — later-in-body marker (gmail) does not win over an earlier one (divRplyFwdMsg)', () => {
    // divRplyFwdMsg (1st in the boundary list) appears EARLIER in the body than
    // gmail_quote (5th). A broken `cut === -1 → true` guard would unconditionally
    // overwrite cut to the LAST matching boundary (gmail, later) — caught here.
    const html = '<p>keep</p><div id="divRplyFwdMsg">A</div><p>mid</p><div class="gmail_quote">B</div>';
    const r = stripQuotedReplies(html);
    expect(r.stripped).toBe(true);
    expect(r.html).toBe('<p>keep</p><p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>');
  });

  it('stripQuotedReplies cuts at the EARLIEST marker — later-in-list boundary (gmail) appearing earlier in the body still wins', () => {
    // gmail_quote (5th in the boundary list) appears BEFORE divRplyFwdMsg (1st) in the body
    const html = '<p>keep</p><div class="gmail_quote">A</div><div id="divRplyFwdMsg">B</div>';
    const r = stripQuotedReplies(html);
    expect(r.stripped).toBe(true);
    expect(r.html).toBe('<p>keep</p><p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>');
  });

  // Surrounding / trailing whitespace at every `\s*` position so the regexes'
  // whitespace matchers are actually exercised (kills the `\s`→`\S` mutants).
  const plainFixtures = [
    { label: 'Outlook Original Message banner', text: 'My reply.\n\n----- Original Message -----  \nFrom: Alice\nold body' },
    { label: 'Outlook underscore rule', text: 'My reply.\n\n________________________________  \nFrom: Alice\nold body' },
    { label: 'Gmail On … wrote attribution', text: 'My reply.\n\nOn Mon, May 5, 2026 at 3:00 PM Alice <alice@x> wrote:  \nold body' },
    { label: 'leading quote line', text: 'My reply.\n\n> old quoted line\n> more quoted' },
  ];
  it.each(plainFixtures)('stripQuotedPlainText cuts a plain-text body at the $label boundary', ({ text }) => {
    const r = stripQuotedPlainText(text);
    expect(r.stripped).toBe(true);
    expect(r.text.startsWith('My reply.')).toBe(true);
    expect(r.text).not.toContain('old');
    expect(r.text).toContain('[Quoted reply chain removed');
  });

  it('stripQuotedPlainText leaves a plain-text body with no quote markers untouched', () => {
    const text = 'Just a normal note.\nSecond line with a > arrow mid-sentence is fine when not line-leading.';
    const r = stripQuotedPlainText(text);
    expect(r.stripped).toBe(false);
    expect(r.text).toBe(text);
  });

  it('stripQuotedPlainText cuts at the EARLIEST plain-text marker when several are present', () => {
    // the leading-`>` line (4th in the list) appears before the "On … wrote:" line (3rd)
    const text = 'keep\n> early quote\nOn Tue Bob wrote:\nlater';
    const r = stripQuotedPlainText(text);
    expect(r.stripped).toBe(true);
    expect(r.text).toBe('keep\n[Quoted reply chain removed — pass --keep-quoted true to include it]');
  });
});

describe('convert-mail-to-markdown — quoted-text stripping', () => {
  const htmlMsg = (content: string, hasAttachments = false): FetchFn => {
    const body = { subject: 'Re: Q3', body: { contentType: 'html', content }, hasAttachments };
    return async (url) => {
      if (url.includes('/attachments')) return Response.json({ value: [{ id: 'att-broken-shape' }], extra: 'forces a list when hasAttachments' });
      return Response.json(body);
    };
  };

  it('strips the quoted reply chain by default and flags it in the note', async () => {
    const content = '<p>My reply.</p><div id="divRplyFwdMsg"><b>From:</b> Alice<br><b>Sent:</b> Friday</div><p>Original message text that should be dropped.</p>';
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), htmlMsg(content)), { messageId: 'm1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string; note?: string };
    expect(v.text).toContain('My reply.');
    expect(v.text).not.toContain('Original message text');
    expect(v.text).toContain('Quoted reply chain removed');
    expect(v.note).toContain('quoted reply chain stripped');
  });

  it('preserves the quoted reply chain with --keep-quoted true and sets no note', async () => {
    const content = '<p>My reply.</p><div id="divRplyFwdMsg"><b>From:</b> Alice</div><p>Original message text.</p>';
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), htmlMsg(content)), { messageId: 'm1', keepQuoted: 'true' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string; note?: string };
    expect(v.text).toContain('Original message text');
    expect(v.text).not.toContain('Quoted reply chain removed');
    expect(v.note).toBeUndefined();
  });

  it('does not treat an HTML marker as a plain-text quote boundary (HTML markers are not plain-text markers)', async () => {
    const content = 'My reply.\n\n<div id="divRplyFwdMsg">not real html</div>';
    const fetchFn: FetchFn = async () => Response.json({ subject: 'Re', body: { contentType: 'text', content }, hasAttachments: false });
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string; note?: string };
    expect(v.text).toContain('divRplyFwdMsg'); // no plain-text marker present → body untouched
    expect(v.note).toBeUndefined();
  });

  it('strips a quoted reply chain from a plain-text body via the plain-text markers', async () => {
    const content = 'My reply here.\n\nOn Mon, May 5, 2026 at 3:00 PM Alice <alice@x> wrote:\nThe original message text to drop.';
    const fetchFn: FetchFn = async () => Response.json({ subject: 'Re: Q3', body: { contentType: 'text', content }, hasAttachments: false });
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string; note?: string };
    expect(v.text).toContain('My reply here.');
    expect(v.text).not.toContain('original message text to drop');
    expect(v.text).toContain('[Quoted reply chain removed');
    expect(v.note).toContain('quoted reply chain stripped');
  });

  it('preserves a plain-text quoted chain with --keep-quoted true', async () => {
    const content = 'My reply.\n\nOn Mon Alice wrote:\nThe original message.';
    const fetchFn: FetchFn = async () => Response.json({ subject: 'Re', body: { contentType: 'text', content }, hasAttachments: false });
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1', keepQuoted: 'true' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string; note?: string };
    expect(v.text).toContain('The original message.');
    expect(v.text).not.toContain('[Quoted reply chain removed');
    expect(v.note).toBeUndefined();
  });

  it('sets no quoted note when an html body has no quote markers', async () => {
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), htmlMsg('<p>standalone reply</p>')), { messageId: 'm1' });
    expect(result?.ok).toBe(true);
    if (result?.ok) expect((result.value as { note?: string }).note).toBeUndefined();
  });

  it('joins the attachments-list note and the quoted-strip note with "; "', async () => {
    // hasAttachments:true → the attachments list is fetched; the malformed shape
    // sets attachmentsListNote, and the quoted block sets the strip note → both joined.
    const content = '<p>reply</p><div id="divRplyFwdMsg">x</div><p>old</p>';
    const fetchFn: FetchFn = async (url) => {
      if (url.includes('/attachments')) return Response.json({ value: 'not-an-array' });
      return Response.json({ subject: 'Re', body: { contentType: 'html', content }, hasAttachments: true });
    };
    const result = await cmdMap['convert-mail-to-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { messageId: 'm1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { note?: string };
    expect(v.note).toContain('quoted reply chain stripped');
    expect(v.note).toContain('malformed shape');
    expect(v.note).toContain('; ');
  });
});

describe('onenote-resource-embedder', () => {
  const RES = 'https://graph.microsoft.com/v1.0/users/u1/onenote/resources/r1/$value';
  // Strict on the exact resource URL so a broken base-strip (wrong getBinary path) → 404 → no embed.
  const imageFetch =
    (contentType: string, bytes: Uint8Array): FetchFn =>
    async (url) =>
      url === RES ? new Response(bytes as unknown as BodyInit, { status: 200, headers: { 'content-type': contentType } }) : new Response('no', { status: 404 });

  it('embeds a OneNote resource image as a base64 data URI (fetching the base-stripped resource path)', async () => {
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), imageFetch('image/png', new Uint8Array([1, 2, 3]))), `<p>see <img src="${RES}" alt="x"></p>`);
    expect(out).toContain('data:image/png;base64,AQID');
    expect(out).not.toContain(RES);
  });

  it('matches the image content-type case-insensitively (IMAGE/PNG embeds)', async () => {
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), imageFetch('IMAGE/PNG', new Uint8Array([1, 2, 3]))), `<img src="${RES}">`);
    expect(out).toContain('data:IMAGE/PNG;base64,AQID');
  });

  it('leaves the URL in place when the resource is not an image', async () => {
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), imageFetch('application/pdf', new Uint8Array([1, 2, 3]))), `<img src="${RES}">`);
    expect(out).toContain(RES);
    expect(out).not.toContain('data:');
  });

  it('embeds at exactly the size limit but leaves the URL one byte over (boundary)', async () => {
    const atLimit = await embedOnenoteResources(createGraphClient(fakeAuth(), imageFetch('image/png', new Uint8Array(2_000_000))), `<img src="${RES}">`);
    expect(atLimit).toContain('data:image/png;base64,'); // 2_000_000 is NOT over the limit → embedded
    expect(atLimit).not.toContain(RES);

    const overLimit = await embedOnenoteResources(createGraphClient(fakeAuth(), imageFetch('image/png', new Uint8Array(2_000_001))), `<img src="${RES}">`);
    expect(overLimit).toContain(RES); // one byte over → left as URL
    expect(overLimit).not.toContain('data:image');
  });

  it('leaves the URL in place when the resource fetch fails', async () => {
    const out = await embedOnenoteResources(
      createGraphClient(fakeAuth(), async () => new Response('nope', { status: 500 })),
      `<img src="${RES}">`
    );
    expect(out).toContain(RES);
  });

  it('leaves the URL in place (no throw) when the resource has no inline bytes — a 302 redirect with no base64/contentType', async () => {
    const redirect = createGraphClient(fakeAuth(), async () => new Response(null, { status: 302, headers: { location: 'https://francecentral1-mediap.svc.ms/x' } }));
    const out = await embedOnenoteResources(redirect, `<img src="${RES}">`);
    expect(out).toContain(RES);
    expect(out).not.toContain('data:');
  });

  it('leaves the URL in place (no throw) when the resource response has base64 but no contentType (JSON envelope)', async () => {
    const jsonNoType = createGraphClient(fakeAuth(), async () => Response.json({ base64: 'AQID' })); // no contentType field
    const out = await embedOnenoteResources(jsonNoType, `<img src="${RES}">`);
    expect(out).toContain(RES);
    expect(out).not.toContain('data:');
  });

  it('fetches each distinct resource once (dedupe) and embeds every occurrence', async () => {
    let calls = 0;
    const graph = createGraphClient(fakeAuth(), async (url) => {
      if (url !== RES) return new Response('no', { status: 404 });
      calls += 1;
      return new Response(new Uint8Array([1, 2, 3]), { status: 200, headers: { 'content-type': 'image/png' } });
    });
    const out = await embedOnenoteResources(graph, `<img src="${RES}"> mid <img src="${RES}">`);
    expect(calls).toBe(1);
    expect(out.match(/data:image\/png;base64,AQID/g) ?? []).toHaveLength(2);
    expect(out).not.toContain(RES);
  });

  it('returns the html unchanged AND fetches nothing when there are no OneNote resource URLs', async () => {
    let calls = 0;
    const graph = createGraphClient(fakeAuth(), async () => {
      calls += 1;
      return new Response('unexpected', { status: 500 });
    });
    const html = '<p>plain page with an <img src="https://example.com/x.png"> external image</p>';
    expect(await embedOnenoteResources(graph, html)).toBe(html);
    expect(calls).toBe(0); // no match → empty array default → no getBinary call (kills the `?? []` mutant)
  });

  // <object> file attachments are annotated (not fetched) — the fetch must NOT fire.
  const noFetch: FetchFn = async () => {
    throw new Error('object attachments must not trigger a fetch');
  };

  it('annotates an <object> file attachment with its name + type and drops the raw element', async () => {
    const html = `<p>doc</p><object data="${RES}" data-attachment="report.pdf" type="application/pdf"></object>`;
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), noFetch), html);
    expect(out).toContain('[OneNote attachment: report.pdf (application/pdf)]');
    expect(out).not.toContain('<object');
    expect(out).not.toContain(RES);
  });

  it('annotates an <object> attachment without a type (no parenthetical)', async () => {
    const html = `<object data="${RES}" data-attachment="notes.docx"></object>`;
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), noFetch), html);
    expect(out).toContain('[OneNote attachment: notes.docx]');
    expect(out).not.toContain('(');
  });

  it('falls back to a generic name for an <object> attachment with no data-attachment', async () => {
    const html = `<object data="${RES}" type="application/pdf"></object>`;
    const out = await embedOnenoteResources(createGraphClient(fakeAuth(), noFetch), html);
    expect(out).toContain('[OneNote attachment: file (application/pdf)]');
  });
});

describe('onenote-metadata', () => {
  it('renders every field (newline-joined, in order) including the expanded parent section + notebook names', () => {
    const md = formatOnenoteMetadata({
      title: 'Q3 Plan',
      createdDateTime: '2026-01-02T03:04:05Z',
      lastModifiedDateTime: '2026-05-06T07:08:09Z',
      parentSection: { displayName: 'Planning' },
      parentNotebook: { displayName: 'Work' },
    });
    // exact output asserts the `\n` join + line ordering (kills the join-separator mutant)
    expect(md).toBe(
      [
        '## OneNote metadata',
        '- **Title:** Q3 Plan',
        '- **Created:** 2026-01-02T03:04:05Z',
        '- **Last modified:** 2026-05-06T07:08:09Z',
        '- **Notebook:** Work',
        '- **Section:** Planning',
      ].join('\n')
    );
  });

  it('omits absent fields and tolerates a missing parentSection / parentNotebook (no throw)', () => {
    const md = formatOnenoteMetadata({ title: 'Only Title' });
    expect(md).toContain('- **Title:** Only Title');
    expect(md).not.toContain('Created');
    expect(md).not.toContain('Notebook');
    expect(md).not.toContain('Section');
  });

  it('treats an empty-string field as absent (no blank line emitted)', () => {
    const md = formatOnenoteMetadata({ title: '', createdDateTime: '2026-01-01T00:00:00Z' });
    expect(md).not.toContain('Title');
    expect(md).toContain('- **Created:** 2026-01-01T00:00:00Z');
  });

  it('returns just the header when the page has no metadata fields', () => {
    expect(formatOnenoteMetadata({})).toBe('## OneNote metadata');
  });
});

describe('get-onenote-page-as-markdown — inline images + metadata', () => {
  const RES = 'https://graph.microsoft.com/v1.0/me/onenote/resources/r1/$value';
  const pageHtml = `<html><body><p>Notes.</p><img src="${RES}"></body></html>`;
  const onenoteFetch = (): FetchFn => async (url) => {
    if (url.endsWith('/onenote/pages/p1/content')) return new Response(pageHtml, { status: 200, headers: { 'content-type': 'text/html' } });
    if (url === RES) return new Response(new Uint8Array([1, 2, 3]), { status: 200, headers: { 'content-type': 'image/png' } });
    if (url.includes('/onenote/pages/p1?')) {
      return Response.json({
        title: 'My Page',
        createdDateTime: '2026-01-01T00:00:00Z',
        lastModifiedDateTime: '2026-02-02T00:00:00Z',
        parentSection: { displayName: 'Sec' },
        parentNotebook: { displayName: 'NB' },
      });
    }
    return new Response('no', { status: 404 });
  };

  it('embeds inline resource images by default', async () => {
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), onenoteFetch()), { onenotePageId: 'p1' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string };
    expect(v.text).toContain('Notes.');
    expect(v.text).toContain('data:image/png;base64,AQID');
    expect(v.text).not.toContain('resources/r1/$value');
  });

  it('keeps the raw resource URL with --inline-images false', async () => {
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), onenoteFetch()), { onenotePageId: 'p1', inlineImages: 'false' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string };
    expect(v.text).toContain('resources/r1/$value');
    expect(v.text).not.toContain('data:image');
  });

  it('embeds inline resource images when --inline-images true is passed explicitly', async () => {
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), onenoteFetch()), { onenotePageId: 'p1', inlineImages: 'true' });
    expect(result?.ok).toBe(true);
    if (result?.ok) expect((result.value as { text: string }).text).toContain('data:image/png;base64,AQID');
  });

  it('appends the OneNote metadata block (after a blank line) with --include-metadata true, expanding section + notebook', async () => {
    let metadataUrl = '';
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/onenote/pages/p1/content')) return new Response('<p>Body.</p>', { status: 200, headers: { 'content-type': 'text/html' } });
      if (url.includes('/onenote/pages/p1?')) {
        metadataUrl = url;
        return Response.json({
          title: 'My Page',
          createdDateTime: '2026-01-01T00:00:00Z',
          lastModifiedDateTime: '2026-02-02T00:00:00Z',
          parentSection: { displayName: 'Sec' },
          parentNotebook: { displayName: 'NB' },
        });
      }
      return new Response('no', { status: 404 });
    };
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { onenotePageId: 'p1', includeMetadata: 'true' });
    expect(result?.ok).toBe(true);
    if (!result?.ok) return;
    const v = result.value as { text: string };
    // metadata GET expands the parent section + notebook display names (kills the query-string mutant)
    expect(metadataUrl).toContain('$select=title');
    expect(metadataUrl).toContain('$expand=parentSection');
    expect(metadataUrl).toContain('parentNotebook');
    // body, then a blank line, then the metadata block (kills the `\n\n` separator mutant)
    expect(v.text).toContain('Body.\n\n## OneNote metadata');
    expect(v.text).toContain('- **Title:** My Page');
    expect(v.text).toContain('- **Notebook:** NB');
    expect(v.text).toContain('- **Section:** Sec');
  });

  it('omits the metadata block by default and when --include-metadata false is passed explicitly', async () => {
    const def = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), onenoteFetch()), { onenotePageId: 'p1' });
    expect(def?.ok).toBe(true);
    if (def?.ok) expect((def.value as { text: string }).text).not.toContain('## OneNote metadata');

    const explicitFalse = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), onenoteFetch()), { onenotePageId: 'p1', includeMetadata: 'false' });
    expect(explicitFalse?.ok).toBe(true);
    if (explicitFalse?.ok) expect((explicitFalse.value as { text: string }).text).not.toContain('## OneNote metadata');
  });

  it('propagates a content-fetch failure', async () => {
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(
      createGraphClient(
        fakeAuth(),
        async () => new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } })
      ),
      { onenotePageId: 'p1' }
    );
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('api_error');
  });

  it('propagates a metadata-fetch failure under --include-metadata true', async () => {
    const fetchFn: FetchFn = async (url) => {
      if (url.endsWith('/onenote/pages/p1/content')) return new Response('<p>x</p>', { status: 200, headers: { 'content-type': 'text/html' } });
      if (url.includes('/onenote/pages/p1?'))
        return new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'gone' } }), { status: 404, headers: { 'content-type': 'application/json' } });
      return new Response('no', { status: 404 });
    };
    const result = await cmdMap['get-onenote-page-as-markdown']?.execute(createGraphClient(fakeAuth(), fetchFn), { onenotePageId: 'p1', includeMetadata: 'true' });
    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.type).toBe('api_error');
  });
});

type CommandFixture = { readonly name: string; readonly params: Record<string, string>; readonly responseBody?: unknown };

const allCommandFixtures: CommandFixture[] = [
  { name: 'list-drives', params: {} },
  { name: 'get-drive-root-item', params: { driveId: 'd1' } },
  { name: 'list-folder-files', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'download-drive-item-content', params: { driveId: 'd1', itemId: 'i1' }, responseBody: { contentType: 'application/octet-stream', size: 5, base64: 'JVBERi0=' } },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' } },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' } },
  {
    name: 'download-drive-item-version',
    params: { driveId: 'd1', itemId: 'i1', versionId: '3.0' },
    responseBody: { contentType: 'application/octet-stream', size: 5, base64: 'JVBERi0=' },
  },
  { name: 'download-drive-item-as-pdf', params: { driveId: 'd1', itemId: 'i1' }, responseBody: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } },
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
  { name: 'list-teams-chats-with-messages', params: {} },
  { name: 'list-teams-chat-messages', params: { chatId: '19:abc@thread.v2' } },
  { name: 'get-teams-chat-message', params: { chatId: '19:abc@thread.v2', messageId: 'm1' } },
  { name: 'resolve-teams-link', params: { url: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000?ctx=chat' } },
  { name: 'resolve-mail-link', params: { url: 'https://outlook.office.com/mail/inbox/id/AAMkAGI2Test' } },
  { name: 'resolve-drive-share-link', params: { url: 'https://contoso.sharepoint.com/:b:/s/team/EaBcDef123' } },
  { name: 'resolve-calendar-link', params: { url: 'https://outlook.office.com/calendar/item/AAMkAGI2Cal' } },
  { name: 'find-chats-with-user', params: { name: 'nobody' } },
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
  { name: 'get-excel-chart-image', params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', chartId: 'c1' }, responseBody: { value: 'aW1n' } },
  { name: 'list-calendar-event-attachments', params: { eventId: 'e1' } },
  {
    name: 'convert-calendar-event-attachment-to-markdown',
    params: { eventId: 'e1', attachmentId: 'a1' },
    responseBody: { '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentBytes: 'aGk=' },
  },
  {
    name: 'convert-calendar-event-attachment-to-pdf',
    params: { eventId: 'e1', attachmentId: 'a1' },
    responseBody: { '@odata.type': '#microsoft.graph.fileAttachment', name: 'note.txt', contentBytes: 'aGk=' },
  },
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
    { name: 'download-drive-item-content', params: { driveId: 'd1' } },
    { name: 'download-drive-item-version', params: { driveId: 'd1', itemId: 'i1' } },
    { name: 'download-drive-item-as-pdf', params: { driveId: 'd1' } },
    { name: 'download-drive-item-as-markdown', params: { driveId: 'd1' } },
    { name: 'get-onenote-page-as-markdown', params: {} },
    { name: 'search-mail-messages', params: {} },
    { name: 'extract-sharepoint-links-in-mail', params: {} },
    { name: 'extract-sharepoint-links-in-documents', params: { driveId: 'd1' } },
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
    { name: 'list-teams-chat-messages', params: {} },
    { name: 'get-teams-chat-message', params: { chatId: '19:abc@thread.v2' } },
    { name: 'resolve-teams-link', params: {} },
    { name: 'resolve-mail-link', params: {} },
    { name: 'resolve-drive-share-link', params: {} },
    { name: 'resolve-calendar-link', params: {} },
    { name: 'find-chats-with-user', params: {} },
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
  { name: 'download-drive-item-content', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/content' },
  { name: 'get-drive-item', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1' },
  { name: 'list-drive-item-permissions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/permissions' },
  { name: 'list-drive-item-versions', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/versions' },
  {
    name: 'download-drive-item-version',
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
    // No-extension metadata response → the dispatcher content-sniffs the bytes, so
    // the LAST URL hit is the /content fetch (sniff fallback), not the metadata GET.
    expectedPath: '/drives/d1/items/i1/content',
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
  {
    name: 'list-mail-messages',
    params: {},
    expectedPath: '/me/messages?$select=id%2Csubject%2Cfrom%2CtoRecipients%2CccRecipients%2CreceivedDateTime%2ChasAttachments%2CisRead%2Cimportance%2CbodyPreview',
  },
  // Audit Alex-session §A: explicit --select wins over the slim default.
  { name: 'list-mail-messages', params: { select: 'id,subject' }, expectedPath: '/me/messages?$select=id%2Csubject' },
  { name: 'list-mail-folders', params: {}, expectedPath: '/me/mailFolders' },
  { name: 'list-mail-child-folders', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/childFolders' },
  { name: 'list-mail-folder-messages', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messages' },
  {
    name: 'get-mail-message',
    params: { messageId: 'm1' },
    expectedPath: '/me/messages/m1?$select=id%2Csubject%2Cfrom%2CtoRecipients%2CccRecipients%2CreceivedDateTime%2ChasAttachments%2CisRead%2Cimportance%2CbodyPreview',
  },
  { name: 'get-mail-message', params: { messageId: 'm1', select: 'id,subject' }, expectedPath: '/me/messages/m1?$select=id%2Csubject' },
  {
    name: 'list-mail-attachments',
    params: { messageId: 'm1' },
    expectedPath: '/me/messages/m1/attachments?$select=id%2Cname%2CcontentType%2Csize%2CisInline',
  },
  { name: 'get-mail-attachment', params: { messageId: 'm1', attachmentId: 'a1' }, expectedPath: '/me/messages/m1/attachments/a1' },
  { name: 'list-mail-rules', params: { mailFolderId: 'f1' }, expectedPath: '/me/mailFolders/f1/messageRules' },
  { name: 'get-mailbox-settings', params: {}, expectedPath: '/me/mailboxSettings' },
  // v1.4.0 fresh-pass #4: search-mail-messages now mirrors list-mail-messages'
  // slim default. The path always carries the inlined $select.
  {
    name: 'search-mail-messages',
    params: { query: 'invoice' },
    expectedPath:
      '/me/messages?$search="invoice"&$select=id%2Csubject%2Cfrom%2CtoRecipients%2CccRecipients%2CreceivedDateTime%2ChasAttachments%2CisRead%2Cimportance%2CbodyPreview',
  },
  // Explicit --select wins over the slim default.
  { name: 'search-mail-messages', params: { query: 'invoice', select: 'id,subject' }, expectedPath: '/me/messages?$search="invoice"&$select=id%2Csubject' },
  { name: 'extract-sharepoint-links-in-mail', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1?$select=subject,body' },
  { name: 'convert-mail-to-markdown', params: { messageId: 'm1' }, expectedPath: '/me/messages/m1' },
  { name: 'list-onenote-notebooks', params: {}, expectedPath: '/me/onenote/notebooks' },
  { name: 'list-onenote-notebook-sections', params: { notebookId: 'n1' }, expectedPath: '/me/onenote/notebooks/n1/sections' },
  { name: 'list-all-onenote-sections', params: {}, expectedPath: '/me/onenote/sections' },
  { name: 'list-onenote-section-pages', params: { onenoteSectionId: 's1' }, expectedPath: '/me/onenote/sections/s1/pages' },
  { name: 'get-onenote-page-content', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'get-onenote-page-as-markdown', params: { onenotePageId: 'p1' }, expectedPath: '/me/onenote/pages/p1/content' },
  { name: 'search-onenote-pages', params: { titleSubstring: 'meeting' }, expectedPath: "/me/onenote/pages?$filter=contains(title,'meeting')" },
  {
    name: 'get-current-user',
    params: {},
    expectedPath: '/me?$select=id%2CdisplayName%2Cmail%2CuserPrincipalName%2CjobTitle%2CofficeLocation%2CmobilePhone',
  },
  { name: 'get-current-user', params: { select: 'id,displayName' }, expectedPath: '/me?$select=id%2CdisplayName' },
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
  { name: 'list-chats', params: {}, expectedPath: '/me/chats?$select=id%2Ctopic%2CchatType%2CcreatedDateTime%2ClastUpdatedDateTime' },
  { name: 'list-chats', params: { select: 'id,topic' }, expectedPath: '/me/chats?$select=id%2Ctopic' },
  { name: 'get-chat', params: { chatId: 'ch1' }, expectedPath: '/chats/ch1?$select=id%2Ctopic%2CchatType%2CcreatedDateTime%2ClastUpdatedDateTime' },
  { name: 'get-chat', params: { chatId: 'ch1', select: 'id,topic' }, expectedPath: '/chats/ch1?$select=id%2Ctopic' },
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
  {
    name: 'get-excel-chart-image',
    params: { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', chartId: 'c1' },
    expectedPath: "/drives/d1/items/i1/workbook/worksheets/Sheet1/charts/c1/Image(width=0,height=0,fittingMode='Fit')",
  },
  { name: 'convert-calendar-event-attachment-to-markdown', params: { eventId: 'e1', attachmentId: 'a1' }, expectedPath: '/me/events/e1/attachments/a1' },
  { name: 'convert-calendar-event-attachment-to-pdf', params: { eventId: 'e1', attachmentId: 'a1' }, expectedPath: '/me/events/e1/attachments/a1' },
  { name: 'convert-drive-item-zip', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/content' },
  { name: 'extract-sharepoint-links-in-documents', params: { driveId: 'd1', itemId: 'i1' }, expectedPath: '/drives/d1/items/i1/content' },
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

// chatsvcagg-tier commands hit a different host than the Graph base URL —
// they sign with the Teams substrate bearer captured at login AND inject the
// regional segment between the host and path. The fake auth above returns
// `region: 'emea'` so every fixture URL lives under `/api/csa/emea/`.
const chatsvcaggPathFixtures: Array<{ name: string; params: Record<string, string>; expectedUrl: string }> = [
  {
    name: 'list-teams-chats-with-messages',
    params: {},
    expectedUrl:
      'https://teams.microsoft.com/api/csa/emea/api/v3/teams/users/me/chats?pageSize=100&enableMembershipSummary=true&supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true&enableEngageCommunities=false',
  },
  {
    name: 'list-teams-chat-messages',
    params: { chatId: '19:abc@thread.v2' },
    expectedUrl: 'https://teams.microsoft.com/api/csa/emea/api/v1/chats/19%3Aabc%40thread.v2/messages',
  },
  {
    name: 'get-teams-chat-message',
    params: { chatId: '19:abc@thread.v2', messageId: '1700000000001' },
    expectedUrl: 'https://teams.microsoft.com/api/csa/emea/api/v1/chats/19%3Aabc%40thread.v2/messages/1700000000001',
  },
];

describe('chatsvcagg-tier commands call the Teams substrate aggregator', () => {
  it.each(chatsvcaggPathFixtures)('$name calls $expectedUrl', async ({ name, params, expectedUrl }) => {
    const url = await capturedUrl(name, params);
    expect(url).toBe(expectedUrl);
  });

  // list-teams-chats-with-messages routes to the paginated `/chats` sibling
  // endpoint (post 2026-05-21 audit — `/teams/users/me` is the aggregate
  // and caps at 273 with no working cursor; `/teams/users/me/chats` is the
  // dedicated paginated chat-list). Mutation guard on the URL shape +
  // fixed substrate query string.
  it('list-teams-chats-with-messages issues the fixed substrate query string + pageSize default', async () => {
    const url = await capturedUrl('list-teams-chats-with-messages', {});
    expect(url).toContain('/teams/users/me/chats?');
    expect(url).toContain('pageSize=100');
    expect(url).toContain('enableMembershipSummary=true');
    expect(url).toContain('supportsAdditionalSystemGeneratedFolders=true');
    expect(url).toContain('supportsSliceItems=true');
    expect(url).toContain('enableEngageCommunities=false');
    // The aggregate endpoint had `isPrefetch=false`; the paginated /chats
    // endpoint does not. Regression guard against accidentally routing
    // back to the aggregate.
    expect(url).not.toContain('isPrefetch');
  });

  it('list-teams-chats-with-messages appends continuationToken when --continuation-token is provided', async () => {
    const url = await capturedUrl('list-teams-chats-with-messages', { continuationToken: 'opaque-cursor-xyz' });
    expect(url).toContain('continuationToken=opaque-cursor-xyz');
  });

  it('list-teams-chats-with-messages does NOT append continuationToken when omitted', async () => {
    const url = await capturedUrl('list-teams-chats-with-messages', {});
    expect(url).not.toContain('continuationToken');
  });

  it('list-teams-chats-with-messages honours --page-size override', async () => {
    const url = await capturedUrl('list-teams-chats-with-messages', { pageSize: '25' });
    expect(url).toContain('pageSize=25');
  });

  it('list-teams-chats-with-messages rejects pageSize "0" (regex anchor guard)', async () => {
    const cmd = cmdMap['list-teams-chats-with-messages'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { pageSize: '0' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  // list-teams-chat-messages issues a bare path with no query string. The
  // chatsvcagg substrate caps at the 200 most recent messages and ignores
  // every cursor/page-size param we've probed — verified 2026-05-21 — so
  // the CLI ships only the chat-id (anything else would be a misleading
  // no-op on the wire).
  it('list-teams-chat-messages issues a bare path with no query string (substrate has no working pagination)', async () => {
    const url = await capturedUrl('list-teams-chat-messages', { chatId: '19:abc@thread.v2' });
    expect(url).not.toContain('?');
    expect(url).not.toContain('pageSize');
    expect(url).not.toContain('messageToken');
  });

  // Chat-ID + Message-ID are URI-encoded into the path (colons +
  // ats are %-encoded). Mutation guard on the encodeURIComponent call.
  it('list-teams-chat-messages URI-encodes the chat id in the path', async () => {
    const url = await capturedUrl('list-teams-chat-messages', { chatId: '19:abc@thread.v2' });
    expect(url).toContain('/chats/19%3Aabc%40thread.v2/messages');
  });

  it('get-teams-chat-message URI-encodes both chat id AND message id in the path', async () => {
    const url = await capturedUrl('get-teams-chat-message', { chatId: '19:abc@thread.v2', messageId: 'm/with/slashes' });
    expect(url).toContain('/chats/19%3Aabc%40thread.v2/messages/m%2Fwith%2Fslashes');
  });
});

// list-teams-chat-history is the IC3-substrate deep-history command — it
// follows the server-provided `_metadata.syncState` URL backward through
// history, capped by `--max-pages`. Sibling of list-teams-chat-messages
// but bypasses the 200-cap (IC3 has a real cursor; chatsvcagg doesn't).
describe('find-chats-with-user paginates the chat-list and filters members', () => {
  // Stateful fakeFetch — returns a different page per call, mimicking the
  // paginated /chats endpoint's continuationToken flow.
  type ChatFix = {
    id: string;
    title?: string | null;
    chatType?: string;
    members: Array<{ mri?: string; displayName?: string; email?: string; userSubType?: string; userPrincipalName?: string; givenName?: string; surname?: string }>;
  };
  type PageFix = { chats: Array<ChatFix>; continuationToken?: string; hasMoreData?: boolean };
  const sequencedFetch = (pages: ReadonlyArray<PageFix>): FakeFetch => {
    let calls = 0;
    let lastUrl: string | null = null;
    const urls: string[] = [];
    const fn = async (url: string): Promise<Response> => {
      lastUrl = url;
      urls.push(url);
      const page = pages[Math.min(calls, pages.length - 1)] ?? { chats: [] };
      calls += 1;
      return new Response(JSON.stringify(page), { headers: { 'content-type': 'application/json' } });
    };
    Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
    Object.defineProperty(fn, 'lastBody', { get: () => null });
    Object.defineProperty(fn, 'urls', { get: () => urls });
    return fn as FakeFetch & { urls: string[] };
  };

  // Branches on host: the substrate chat-list walk (teams.microsoft.com) vs the
  // elevated per-chat member hydration (graph.microsoft.com/.../chats/{id}/members).
  // Lets a test drive the cross-tenant path end to end — the summary roster left
  // the counterpart bare, and `/chats/{id}/members` resolves their name/email.
  type GraphMemberFix = { displayName?: string; email?: string; userId?: string };
  type MembersFix = Record<string, { ok: boolean; members?: Array<GraphMemberFix> }>;
  const hydratingFetch = (pages: ReadonlyArray<PageFix>, membersByChat: MembersFix): FakeFetch & { urls: string[] } => {
    let calls = 0;
    let lastUrl: string | null = null;
    const urls: string[] = [];
    const fn = async (url: string): Promise<Response> => {
      lastUrl = url;
      urls.push(url);
      const members = /graph\.microsoft\.com.*\/chats\/([^?]+?)\/members/.exec(url);
      if (members) {
        const entry = membersByChat[decodeURIComponent(members[1] ?? '')];
        if (entry === undefined || !entry.ok)
          return new Response(JSON.stringify({ error: { code: 'Forbidden' } }), { status: 403, headers: { 'content-type': 'application/json' } });
        return new Response(JSON.stringify({ value: entry.members ?? [] }), { headers: { 'content-type': 'application/json' } });
      }
      const page = pages[Math.min(calls, pages.length - 1)] ?? { chats: [] };
      calls += 1;
      return new Response(JSON.stringify(page), { headers: { 'content-type': 'application/json' } });
    };
    Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
    Object.defineProperty(fn, 'lastBody', { get: () => null });
    Object.defineProperty(fn, 'urls', { get: () => urls });
    return fn as FakeFetch & { urls: string[] };
  };

  it('returns chats whose members match the query on displayName (case-insensitive)', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [
          {
            id: '19:a@unq.gbl.spaces',
            title: 'A chat',
            chatType: 'oneOnOne',
            members: [{ mri: '8:orgid:1', displayName: 'Alex Kim', email: 'alex.kim@example.com', userSubType: 'Member' }],
          },
          {
            id: '19:b@unq.gbl.spaces',
            title: 'B chat',
            chatType: 'oneOnOne',
            members: [{ mri: '8:orgid:2', displayName: 'Someone Else', email: 'someone.else@corp.com', userSubType: 'Member' }],
          },
        ],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'alex' }); // lowercase, accent preserved
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matches: Array<{ chatId: string }>; matchCount: number };
      expect(v.matchCount).toBe(1);
      expect(v.matches[0]?.chatId).toBe('19:a@unq.gbl.spaces');
    }
  });

  it('surfaces dual-identity matches — same person, two MRIs across two chats', async () => {
    // Alex has an org MRI in one chat and a guest MRI in another. Both
    // hits surface in `matches`. This is the canonical workflow win.
    const fetchFn = sequencedFetch([
      {
        chats: [
          {
            id: '19:org@unq.gbl.spaces',
            chatType: 'oneOnOne',
            members: [{ mri: '8:orgid:42c44e51', displayName: 'Alex Kim', email: 'alex.kim@example.com', userSubType: 'Member' }],
          },
          {
            id: '19:guest@unq.gbl.spaces',
            chatType: 'oneOnOne',
            members: [{ mri: '8:orgid:a1bb71fb', displayName: 'Alex Kim (Guest)', email: 'alex.kim@partner.example.com', userSubType: 'Guest' }],
          },
        ],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'alex doe' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; matches: Array<{ chatId: string; matchedMembers: Array<{ mri?: string; userSubType?: string }> }> };
      expect(v.matchCount).toBe(2);
      const mris = v.matches.flatMap((m) => m.matchedMembers.map((mm) => mm.mri));
      expect(mris).toContain('8:orgid:42c44e51');
      expect(mris).toContain('8:orgid:a1bb71fb');
      // Surfaces the userSubType so the caller can distinguish org from guest.
      const subtypes = v.matches.flatMap((m) => m.matchedMembers.map((mm) => mm.userSubType));
      expect(subtypes).toContain('Member');
      expect(subtypes).toContain('Guest');
    }
  });

  // Audit Alex-session §D follow-up — the live failure mode that the
  // synthetic dual-identity test above DID NOT cover: the corporate-MRI
  // member entry's `displayName` is the email itself (no accent), while
  // the guest-MRI entry's displayName carries the accented form. A search
  // for `Alex` used to return ONLY the guest chat because `é` and `e`
  // are different bytes — the corporate 1:1 (843 chats deep) was invisible.
  // Diacritic-folding both sides at compare time fixes it.
  it('surfaces dual-identity when the corporate identity stores its displayName as the un-accented email and the guest identity stores the accented name (Audit Alex-session §D)', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [
          {
            id: '19:42c44e51-corp@unq.gbl.spaces',
            chatType: 'oneOnOne',
            // The live shape that fooled the pre-fix predicate: displayName
            // is literally the email, no "Alex" anywhere.
            members: [{ mri: '8:orgid:42c44e51-aaaa-bbbb-cccc-dddddddddddd', displayName: 'alex.kim@example.com', email: 'alex.kim@example.com', userSubType: 'Member' }],
          },
          {
            id: '19:a1bb71fb-guest@unq.gbl.spaces',
            chatType: 'oneOnOne',
            members: [{ mri: '8:orgid:a1bb71fb-aaaa-bbbb-cccc-dddddddddddd', displayName: 'Alex Kim', email: 'alex.kim@external.example.com', userSubType: 'Guest' }],
          },
        ],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Alex' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; matches: Array<{ chatId: string; matchedMembers: Array<{ mri?: string; userSubType?: string }> }> };
      expect(v.matchCount).toBe(2);
      const mris = v.matches.flatMap((m) => m.matchedMembers.map((mm) => mm.mri));
      expect(mris).toContain('8:orgid:42c44e51-aaaa-bbbb-cccc-dddddddddddd');
      expect(mris).toContain('8:orgid:a1bb71fb-aaaa-bbbb-cccc-dddddddddddd');
    }
  });

  it('the inverse of the §D fix: an un-accented query (`Alex`) ALSO surfaces a member whose only matchable text is the accented displayName (`Alex Kim`)', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [
          {
            id: '19:guest@unq.gbl.spaces',
            chatType: 'oneOnOne',
            // Email and MRI deliberately don't contain "alex" — only the
            // accented displayName does, so the fold has to work on the
            // candidate side as well as the query side.
            members: [{ mri: '8:orgid:zz-no-match-in-mri', displayName: 'Alex Kim', email: 'j.d@external.example.com', userSubType: 'Guest' }],
          },
        ],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Alex' });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.value as { matchCount: number }).matchCount).toBe(1);
  });

  it('folding handles other Latin-1 diacritics too (ç ↔ c, ñ ↔ n) — regression guard against narrowly fixing only `é`', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [
          { id: '19:ç@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '8:orgid:1', displayName: 'François Niçoise' }] },
          { id: '19:ñ@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '8:orgid:2', displayName: 'Niño Pequeño' }] },
        ],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const francois = await cmd.execute(graph, { name: 'francois nicoise' });
    expect(francois.ok).toBe(true);
    if (francois.ok) expect((francois.value as { matchCount: number }).matchCount).toBe(1);
    const nino = await cmd.execute(graph, { name: 'nino' });
    expect(nino.ok).toBe(true);
    if (nino.ok) expect((nino.value as { matchCount: number }).matchCount).toBe(1);
  });

  it('matches on email even when display-name is absent (anonymized/system entries)', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [{ id: '19:c@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '8:orgid:3', email: 'alex.kim@corp.com' }] }],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'alex.kim' });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.value as { matchCount: number }).matchCount).toBe(1);
  });

  it('matches on MRI directly (useful when the caller already has the MRI from another command)', async () => {
    const fetchFn = sequencedFetch([
      {
        chats: [{ id: '19:d@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '8:orgid:abc-1234', displayName: 'X Y', email: 'x.y@corp.com' }] }],
      },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'abc-1234' });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.value as { matchCount: number }).matchCount).toBe(1);
  });

  it('walks continuationToken across pages until hasMoreData is false', async () => {
    const fetchFn = sequencedFetch([
      { chats: [{ id: '19:p1@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '1', displayName: 'Alice' }] }], continuationToken: 'cursor-2', hasMoreData: true },
      { chats: [{ id: '19:p2@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '2', displayName: 'Alice Smith' }] }], continuationToken: 'cursor-3', hasMoreData: true },
      { chats: [{ id: '19:p3@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '3', displayName: 'Alice Jones' }] }], hasMoreData: false }, // last page
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'alice' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; pagesFetched: number; hasMore: boolean };
      expect(v.matchCount).toBe(3);
      expect(v.pagesFetched).toBe(3);
      expect(v.hasMore).toBe(false);
    }
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls).toHaveLength(3);
    expect(urls[1]).toContain('continuationToken=cursor-2');
    expect(urls[2]).toContain('continuationToken=cursor-3');
  });

  it('stops at --max-pages and sets hasMore:true + surfaces nextContinuationToken when the server keeps offering more', async () => {
    const fullPage = (cursor: string): PageFix => ({
      chats: [{ id: `19:${cursor}@unq.gbl.spaces`, chatType: 'oneOnOne', members: [{ mri: '0', displayName: 'Match Me' }] }],
      continuationToken: cursor,
      hasMoreData: true,
    });
    const fetchFn = sequencedFetch([fullPage('c2'), fullPage('c3'), fullPage('c4')]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'match', maxPages: '2' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { hasMore: boolean; nextContinuationToken?: string; pagesFetched: number };
      expect(v.hasMore).toBe(true);
      expect(v.nextContinuationToken).toBe('c3');
      expect(v.pagesFetched).toBe(2);
    }
  });

  it('returns an empty match list when no member matches the query', async () => {
    const fetchFn = sequencedFetch([{ chats: [{ id: '19:e@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '8:orgid:9', displayName: 'Bob' }] }] }]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'nonexistent person' });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.value as { matchCount: number; chatsScanned: number }).matchCount).toBe(0);
  });

  it('deduplicates a chat that somehow appears twice across pages (defense in depth)', async () => {
    const fetchFn = sequencedFetch([
      { chats: [{ id: '19:dup@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '1', displayName: 'Dup' }] }], continuationToken: 'c2', hasMoreData: true },
      // Same id again on page 2 — should not double-count.
      { chats: [{ id: '19:dup@unq.gbl.spaces', chatType: 'oneOnOne', members: [{ mri: '1', displayName: 'Dup' }] }], hasMoreData: false },
    ]);
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'dup' });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.value as { matchCount: number }).matchCount).toBe(1);
  });

  it('rejects an empty --name before any HTTP call', async () => {
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { name: '' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  // ── Cross-tenant hydration (the Chen Robin repro): an externally-homed
  //    counterpart returns in the summary roster as a bare `{mri}` — no name,
  //    no email — so a name search can't match it. `/chats/{id}/members`
  //    (elevated) resolves the member; the matcher re-runs on the hydrated
  //    roster. The bug was a silent `matchCount:0` on a real, active chat.
  const robinChenChatId = '19:aaaaaaaa-1111-2222-3333-444444444444_bbbbbbbb-1111-2222-3333-444444444444@unq.gbl.spaces';
  const robinChenMri = '8:orgid:aaaaaaaa-1111-2222-3333-444444444444';

  it('discovers a cross-tenant 1:1 by name when the summary roster left the counterpart bare — hydrates via /chats/{id}/members (Chen Robin repro)', async () => {
    const fetchFn = hydratingFetch([{ chats: [{ id: robinChenChatId, chatType: 'oneOnOne', members: [{ mri: robinChenMri }] }] }], {
      [robinChenChatId]: {
        ok: true,
        members: [
          { displayName: 'Robin Chen', email: 'robin.chen@contoso.com', userId: 'aaaaaaaa-1111-2222-3333-444444444444' },
          { displayName: 'Me', userId: 'bbbbbbbb-1111-2222-3333-444444444444' },
        ],
      },
    });
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Robin Chen' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as {
        matchCount: number;
        chatsHydrated: number;
        unresolvedMemberCount: number;
        hint?: string;
        matches: Array<{ chatId: string; matchedMembers: Array<{ mri?: string; displayName?: string; email?: string }> }>;
      };
      expect(v.matchCount).toBe(1);
      expect(v.matches[0]?.chatId).toBe(robinChenChatId);
      expect(v.matches[0]?.matchedMembers[0]?.displayName).toBe('Robin Chen');
      expect(v.matches[0]?.matchedMembers[0]?.email).toBe('robin.chen@contoso.com');
      // Reconstructs the substrate MRI from the resolved object-id so the caller
      // can feed it straight back to get-chat / list-chat-members.
      expect(v.matches[0]?.matchedMembers[0]?.mri).toBe(robinChenMri);
      expect(v.chatsHydrated).toBe(1);
      expect(v.unresolvedMemberCount).toBe(0);
      expect(v.hint).toBeUndefined();
    }
  });

  it('surfaces a counterpart who is RESOLVED in a meeting but BARE in the 1:1 — hydrates the direct chat even though the cheap pass already matched (the live dual-identity gap)', async () => {
    // The live failure: the cheap pass found Yu in a small meeting (resolved as
    // a Guest, a DIFFERENT object-id), so a "hydrate only when empty" gate never
    // probed the 1:1 where her HOME identity is bare. Both chats must come back.
    const meetingId = '19:meeting_ProjectGroup@thread.v2';
    const fetchFn = hydratingFetch(
      [
        {
          chats: [
            {
              id: meetingId,
              chatType: 'meeting',
              members: [{ mri: '8:orgid:dddddddd-1111-2222-3333-444444444444', displayName: 'Robin Chen', email: 'Robin.chen@contoso.com', userSubType: 'Guest' }],
            },
            { id: robinChenChatId, chatType: 'oneOnOne', members: [{ mri: robinChenMri }] },
          ],
        },
      ],
      { [robinChenChatId]: { ok: true, members: [{ displayName: 'Robin Chen', email: 'robin.chen@contoso.com', userId: 'aaaaaaaa-1111-2222-3333-444444444444' }] } }
    );
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Robin Chen' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; chatsHydrated: number; matches: Array<{ chatId: string }> };
      expect(v.matchCount).toBe(2);
      const ids = v.matches.map((m) => m.chatId);
      expect(ids).toContain(meetingId); // resolved in the cheap pass
      expect(ids).toContain(robinChenChatId); // resolved by hydrating the bare 1:1
      expect(v.chatsHydrated).toBe(1); // only the bare 1:1 — the meeting carried a resolved member
    }
  });

  it('turns a confident empty into an honest one — when hydration cannot resolve a bare member it returns a hint + unresolvedMemberCount, not a bare matchCount:0', async () => {
    const fetchFn = hydratingFetch(
      [{ chats: [{ id: robinChenChatId, chatType: 'oneOnOne', members: [{ mri: robinChenMri }] }] }],
      { [robinChenChatId]: { ok: false } } // /chats/{id}/members → 403
    );
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Robin Chen' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; chatsHydrated: number; unresolvedMemberCount: number; hint?: string };
      expect(v.matchCount).toBe(0);
      expect(v.chatsHydrated).toBe(1);
      expect(v.unresolvedMemberCount).toBe(1);
      expect(typeof v.hint).toBe('string');
      expect(v.hint).toContain('object-id');
    }
  });

  it('hydration that resolves the roster but still does not match yields a clean empty (no hint, unresolvedMemberCount 0) — the person genuinely is not in the chat', async () => {
    const fetchFn = hydratingFetch([{ chats: [{ id: robinChenChatId, chatType: 'oneOnOne', members: [{ mri: robinChenMri }] }] }], {
      [robinChenChatId]: { ok: true, members: [{ displayName: 'Robin Chen', email: 'robin.chen@contoso.com', userId: 'aaaaaaaa-1111-2222-3333-444444444444' }] },
    });
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Nonexistent Person' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; chatsHydrated: number; unresolvedMemberCount: number; hint?: string };
      expect(v.matchCount).toBe(0);
      expect(v.chatsHydrated).toBe(1);
      expect(v.unresolvedMemberCount).toBe(0);
      expect(v.hint).toBeUndefined();
    }
  });

  it('does NOT hydrate bare members in group/meeting chats (cost control) — it counts them in unresolvedMemberCount and, on an empty result, emits the hint', async () => {
    const fetchFn = hydratingFetch(
      [{ chats: [{ id: '19:meeting_big@thread.v2', chatType: 'meeting', members: [{ mri: '8:orgid:aaaaaaaa-1111-2222-3333-444444444444' }, { displayName: 'Me' }] }] }],
      {} // no /chats/{id}/members ever called — assert that below
    );
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Robin Chen' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; chatsHydrated: number; unresolvedMemberCount: number; hint?: string };
      expect(v.matchCount).toBe(0);
      expect(v.chatsHydrated).toBe(0); // the meeting is not a direct chat → not probed
      expect(v.unresolvedMemberCount).toBe(1); // the one bare member is surfaced
      expect(typeof v.hint).toBe('string');
    }
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls.some((u) => u.includes('/members'))).toBe(false);
  });

  it('never hydrates a 1:1 whose member is already name-resolved — a non-matching but resolvable member is not bare, so no per-chat call and unresolvedMemberCount stays 0 (cost-control contract)', async () => {
    // Guards the bareness predicate: an unmatched chat whose members all carry a
    // name/email must NOT be treated as a hydration candidate. Otherwise every
    // search would issue a members lookup for every unrelated 1:1.
    const fetchFn = hydratingFetch(
      [{ chats: [{ id: '19:alice_me@unq.gbl.spaces', chatType: 'chat', members: [{ mri: '8:orgid:alice', displayName: 'Alice Wonder', email: 'alice@corp.com' }] }] }],
      {}
    );
    const cmd = cmdMap['find-chats-with-user'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const r = await cmd.execute(graph, { name: 'Robin Chen' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { matchCount: number; chatsHydrated: number; unresolvedMemberCount: number; hint?: string };
      expect(v.matchCount).toBe(0);
      expect(v.chatsHydrated).toBe(0);
      expect(v.unresolvedMemberCount).toBe(0);
      expect(v.hint).toBeUndefined();
    }
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls.some((u) => u.includes('/members'))).toBe(false);
  });
});

describe('resolve-teams-link parses copy-link URLs into structured ids', () => {
  it('extracts chatId + messageId from a typical Teams copy-link URL', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), {
      url: 'https://teams.microsoft.com/l/message/19%3Abbbbbbbb-1111-2222-3333-444444444444_cccccccc-1111-2222-3333-444444444444%40unq.gbl.spaces/1752206983412?tenantId=tenant-abc&groupId=group-xyz&ctx=chat',
    });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { chatId: string; messageId: string; tenantId?: string; groupId?: string; context?: string };
      expect(v.chatId).toBe('19:bbbbbbbb-1111-2222-3333-444444444444_cccccccc-1111-2222-3333-444444444444@unq.gbl.spaces');
      expect(v.messageId).toBe('1752206983412');
      expect(v.tenantId).toBe('tenant-abc');
      expect(v.groupId).toBe('group-xyz');
      expect(v.context).toBe('chat');
    }
  });

  it('handles URLs with no query string (only the two path segments)', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000' });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { chatId: string; messageId: string; tenantId?: string };
      expect(v.chatId).toBe('19:abc@thread.v2');
      expect(v.messageId).toBe('1700000000000');
      expect(v.tenantId).toBeUndefined();
    }
  });

  it('surfaces optional parentMessageId when the link carries a reply context', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), {
      url: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000001?parentMessageId=1700000000000',
    });
    expect(r.ok).toBe(true);
    if (r.ok) {
      const v = r.value as { parentMessageId?: string };
      expect(v.parentMessageId).toBe('1700000000000');
    }
  });

  it('rejects URLs that are not Teams message links with a clear validation_error', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://example.com/some/other/path' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.message).toContain('not a Teams message link');
    }
  });

  it('rejects a Teams /l/message URL that is missing the message-id segment', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  it('rejects a non-URL string before parsing kicks in', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'not-a-url' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  // v1.4.0 re-audit Nit 1 — three cross-resolver pointers for the
  // teams resolver. Drive-share covers both `*.sharepoint.com` and
  // `1drv.ms` hosts (one detector handles both). Outlook URLs split
  // into mail vs calendar for the tightest cross-pointer.
  it('rejects a SharePoint share URL with a structured cli_reject_drive_share_link_on_teams_resolver code (Nit 1 cross-pointer)', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://contoso.sharepoint.com/:b:/s/team/EaBcDef123_xyz' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.code).toBe('cli_reject_drive_share_link_on_teams_resolver');
      expect(r.error.message).toContain('OneDrive / SharePoint');
    }
  });

  it('also rejects a 1drv.ms short-link URL with the same drive-share cross-pointer code (one detector covers both *.sharepoint.com and 1drv.ms)', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://1drv.ms/b/s!AbCdEfGh_xyz' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.code).toBe('cli_reject_drive_share_link_on_teams_resolver');
    }
  });

  it('rejects an Outlook mail URL with a structured cli_reject_mail_link_on_teams_resolver code (Nit 1 cross-pointer; splits mail vs calendar)', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://outlook.office.com/mail/inbox/id/AAMkAGI2_mail_on_teams' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.code).toBe('cli_reject_mail_link_on_teams_resolver');
    }
  });

  it('rejects an Outlook calendar URL with a structured cli_reject_calendar_link_on_teams_resolver code (Nit 1 cross-pointer)', async () => {
    const cmd = cmdMap['resolve-teams-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url: 'https://outlook.office.com/calendar/item/AAMkAGI2_calendar_on_teams' });
    expect(r.ok).toBe(false);
    if (!r.ok && r.error.type === 'validation_error') {
      expect(r.error.code).toBe('cli_reject_calendar_link_on_teams_resolver');
    }
  });
});

// Outlook web mail links → messageId. Mirrors resolve-teams-link's shape;
// rejects calendar URLs with a pointer to resolve-calendar-link so an LLM
// never silently treats a calendar invite as a mail message.
describe('resolve-mail-link parses Outlook web mail URLs into messageId', () => {
  const exec = async (url: string): Promise<{ ok: boolean; value?: { messageId: string }; error?: GraphError }> => {
    const cmd = cmdMap['resolve-mail-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url });
    return r.ok ? { ok: true, value: r.value as { messageId: string } } : { ok: false, error: r.error };
  };

  it('extracts messageId from the modern path-style URL (`/mail/<folder>/id/<id>`)', async () => {
    const r = await exec('https://outlook.office.com/mail/inbox/id/AAMkAGI2THVS_modern_path');
    expect(r.ok).toBe(true);
    expect(r.value?.messageId).toBe('AAMkAGI2THVS_modern_path');
  });

  it('extracts messageId from the OWA query-style URL with lowercase `itemid`', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMkAGI2THVS_owa_lower&exvsurl=1&path=/mail/inbox');
    expect(r.ok).toBe(true);
    expect(r.value?.messageId).toBe('AAMkAGI2THVS_owa_lower');
  });

  it('extracts messageId from the OWA query-style URL with capitalised `ItemID` (legacy emit)', async () => {
    const r = await exec('https://outlook.office365.com/owa/?ItemID=AAMkAGI2THVS_owa_capital&exvsurl=1&viewmodel=ReadMessageItem');
    expect(r.ok).toBe(true);
    expect(r.value?.messageId).toBe('AAMkAGI2THVS_owa_capital');
  });

  it('URL-decodes the extracted id so percent-encoded ids round-trip correctly', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMk%2BFoo%2FBar%3D');
    expect(r.ok).toBe(true);
    expect(r.value?.messageId).toBe('AAMk+Foo/Bar=');
  });

  it('rejects a calendar OWA link (path=/calendar/item) with a structured code pointing at resolve-calendar-link', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMkAGI2_calendar&exvsurl=1&path=/calendar/item');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_calendar_link_on_mail_resolver');
    expect(r.error?.message).toContain('calendar');
  });

  it('rejects a calendar path-style link (`/calendar/item/...`) with the same structured code', async () => {
    const r = await exec('https://outlook.office.com/calendar/item/AAMkAGI2_calendar_path');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_calendar_link_on_mail_resolver');
  });

  it('rejects an unrelated host (example.com) with a generic "not an Outlook mail link" message', async () => {
    const r = await exec('https://example.com/mail/AAMkAGI2');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('not an Outlook mail link');
  });

  it('rejects a non-URL string at the Zod layer before parsing runs', async () => {
    const r = await exec('not-a-url');
    expect(r.ok).toBe(false);
    expect(r.error?.type).toBe('validation_error');
  });

  it('rejects an outlook.office.com URL with no extractable id (e.g. a bare `/mail/inbox` folder URL)', async () => {
    const r = await exec('https://outlook.office.com/mail/inbox');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('not an Outlook mail link');
  });

  // v1.4.0 re-audit Nit 1 — cross-resolver pointer for the teams case.
  // A Teams /l/message link wrongly passed used to fall through to the
  // generic "not an Outlook mail link" rejection; now emits a structured
  // code the hint table maps to "Re-run with resolve-teams-link".
  it('rejects a Teams `/l/message/...` URL with a structured cli_reject_teams_link_on_mail_resolver code (Nit 1 cross-pointer)', async () => {
    const r = await exec('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_teams_link_on_mail_resolver');
    expect(r.error?.message).toContain('Teams message link');
  });
});

// SharePoint / OneDrive sharing URLs → Graph share token (`u!<base64url>`).
// Pure encoding via the existing `buildShareToken` helper; the test only
// needs to verify the round-trip + the accepted-host gate.
describe('resolve-drive-share-link encodes sharing URLs into Graph /shares/{token}', () => {
  const exec = async (url: string): Promise<{ ok: boolean; value?: { shareToken: string; graphPath: string; originalUrl: string }; error?: GraphError }> => {
    const cmd = cmdMap['resolve-drive-share-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url });
    return r.ok ? { ok: true, value: r.value as { shareToken: string; graphPath: string; originalUrl: string } } : { ok: false, error: r.error };
  };

  it('encodes a tenant SharePoint share URL as `u!<base64url>` and emits the `/shares/{token}/driveItem` Graph path ready for the next call', async () => {
    const url = 'https://contoso.sharepoint.com/:b:/s/team/EaBcDef123_xyz';
    const r = await exec(url);
    expect(r.ok).toBe(true);
    expect(r.value?.shareToken.startsWith('u!')).toBe(true);
    expect(r.value?.graphPath).toBe(`/shares/${r.value?.shareToken}/driveItem`);
    expect(r.value?.originalUrl).toBe(url);
  });

  it('accepts a personal OneDrive (-my.sharepoint.com subdomain) URL via the .sharepoint.com suffix match', async () => {
    const url = 'https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/report.pdf';
    const r = await exec(url);
    expect(r.ok).toBe(true);
    expect(r.value?.shareToken.startsWith('u!')).toBe(true);
  });

  it('accepts a 1drv.ms short-link URL', async () => {
    const url = 'https://1drv.ms/b/s!AbCdEfGh_xyz';
    const r = await exec(url);
    expect(r.ok).toBe(true);
    expect(r.value?.shareToken.startsWith('u!')).toBe(true);
  });

  it('rejects an unrelated host (graph.microsoft.com) with a clear validation_error', async () => {
    const r = await exec('https://graph.microsoft.com/v1.0/me/drive');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('not a recognised OneDrive / SharePoint sharing URL');
  });

  it('rejects a non-URL string at the Zod layer', async () => {
    const r = await exec('not-a-url');
    expect(r.ok).toBe(false);
    expect(r.error?.type).toBe('validation_error');
  });

  it('the encoded shareToken is base64url (no `+`, `/`, or `=` padding) per the Graph /shares/{token} contract', async () => {
    const url = 'https://contoso.sharepoint.com/:b:/s/team/with+plus/and/slash?e=stuff';
    const r = await exec(url);
    expect(r.ok).toBe(true);
    const token = r.value?.shareToken ?? '';
    expect(token.startsWith('u!')).toBe(true);
    expect(token.slice(2)).not.toContain('+');
    expect(token.slice(2)).not.toContain('/');
    expect(token.slice(2)).not.toContain('=');
  });

  // v1.4.0 re-audit Nit 1 — three cross-resolver pointers for the
  // drive-share resolver. Outlook URLs split into mail vs calendar so
  // the cross-pointer goes to the tightest sibling.
  it('rejects an Outlook mail URL with a structured cli_reject_mail_link_on_drive_share_resolver code (Nit 1 cross-pointer)', async () => {
    const r = await exec('https://outlook.office.com/mail/inbox/id/AAMkAGI2_mail_on_drive');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_mail_link_on_drive_share_resolver');
    expect(r.error?.message).toContain('Outlook mail');
  });

  it('rejects an Outlook calendar URL with a structured cli_reject_calendar_link_on_drive_share_resolver code (Nit 1 cross-pointer; splits mail vs calendar so the LLM gets the tightest sibling)', async () => {
    const r = await exec('https://outlook.office.com/calendar/item/AAMkAGI2_calendar_on_drive');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_calendar_link_on_drive_share_resolver');
    expect(r.error?.message).toContain('Outlook calendar');
  });

  it('rejects a Teams `/l/message/...` URL with a structured cli_reject_teams_link_on_drive_share_resolver code (Nit 1 cross-pointer)', async () => {
    const r = await exec('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_teams_link_on_drive_share_resolver');
    expect(r.error?.message).toContain('Teams message link');
  });
});

// Outlook calendar item URLs → eventId. Mirrors resolve-mail-link's shape
// inversely — rejects mail URLs with a pointer to resolve-mail-link.
describe('resolve-calendar-link parses Outlook calendar item URLs into eventId', () => {
  const exec = async (url: string): Promise<{ ok: boolean; value?: { eventId: string }; error?: GraphError }> => {
    const cmd = cmdMap['resolve-calendar-link'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { url });
    return r.ok ? { ok: true, value: r.value as { eventId: string } } : { ok: false, error: r.error };
  };

  it('extracts eventId from the modern path-style URL (`/calendar/item/<id>`)', async () => {
    const r = await exec('https://outlook.office.com/calendar/item/AAMkAGI2CalPath');
    expect(r.ok).toBe(true);
    expect(r.value?.eventId).toBe('AAMkAGI2CalPath');
  });

  it('extracts eventId from the OWA query-style calendar URL (`/owa/?itemid=...&path=/calendar/item`)', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMkAGI2CalOwa&exvsurl=1&path=/calendar/item');
    expect(r.ok).toBe(true);
    expect(r.value?.eventId).toBe('AAMkAGI2CalOwa');
  });

  it('URL-decodes the extracted id so percent-encoded ids round-trip correctly', async () => {
    const r = await exec('https://outlook.office.com/calendar/item/AAMk%2BCal%3D');
    expect(r.ok).toBe(true);
    expect(r.value?.eventId).toBe('AAMk+Cal=');
  });

  it('rejects a mail OWA link (no `path=/calendar`) with a structured code pointing at resolve-mail-link', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMkAGI2Mail&exvsurl=1&viewmodel=ReadMessageItem');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_mail_link_on_calendar_resolver');
    expect(r.error?.message).toContain('mail');
  });

  it('rejects a mail path-style link (`/mail/...`) with the same structured code', async () => {
    const r = await exec('https://outlook.office.com/mail/inbox/id/AAMkAGI2Mail');
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('cli_reject_mail_link_on_calendar_resolver');
  });

  it('rejects an unrelated host with a generic "not an Outlook calendar item link" message', async () => {
    const r = await exec('https://example.com/calendar/item/AAMkAGI2');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('not an Outlook calendar item link');
  });

  it('rejects a non-URL string at the Zod layer', async () => {
    const r = await exec('not-a-url');
    expect(r.ok).toBe(false);
    expect(r.error?.type).toBe('validation_error');
  });

  it('falls through to the generic unknown-link error for an OWA URL whose `path` query is neither `/mail` nor `/calendar` (e.g. `path=/people` — covers the isMailLink negative branch)', async () => {
    const r = await exec('https://outlook.office.com/owa/?itemid=AAMkAGI2_other&path=/people');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('not an Outlook calendar item link');
    expect(r.error?.code).toBeUndefined();
  });
});

describe('list-teams-chat-history follows syncState URLs through IC3 history', () => {
  // Stateful fakeFetch — returns a different page per call, mimicking the
  // IC3 server's response sequence. Tracks every URL hit so tests can
  // assert on the cursor flow.
  type Page = { messages: Array<{ id: string; sequenceId: number }>; _metadata?: { syncState?: string } };
  const sequencedFetch = (pages: ReadonlyArray<Page>): FakeFetch => {
    let calls = 0;
    let lastUrl: string | null = null;
    const urls: string[] = [];
    const fn = async (url: string): Promise<Response> => {
      lastUrl = url;
      urls.push(url);
      const page = pages[Math.min(calls, pages.length - 1)] ?? { messages: [] };
      calls += 1;
      return new Response(JSON.stringify(page), { headers: { 'content-type': 'application/json' } });
    };
    Object.defineProperty(fn, 'lastUrl', { get: () => lastUrl });
    Object.defineProperty(fn, 'lastBody', { get: () => null });
    Object.defineProperty(fn, 'urls', { get: () => urls });
    return fn as FakeFetch & { urls: string[] };
  };

  // Helper: build a syncState absolute URL (server-emitted shape) pointing
  // at the next page. The use-case strips the host+region prefix and uses
  // only the relative path on the next call.
  const syncUrl = (cursor: string): string =>
    `https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/X/messages?startTime=1&syncState=${cursor}&pageSize=100&view=msnp24Equivalent`;

  it('returns a single page and hasMore:false when the first response has no _metadata.syncState (end of history)', async () => {
    const fetchFn = sequencedFetch([{ messages: [{ id: 'm1', sequenceId: 1 }] }]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { chatId: '19:abc@thread.v2' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { messages: unknown[]; hasMore: boolean; nextSyncState?: string };
      expect(v.messages).toHaveLength(1);
      expect(v.hasMore).toBe(false);
      expect(v.nextSyncState).toBeUndefined();
    }
  });

  it('follows _metadata.syncState across pages until the server stops emitting one (end of history)', async () => {
    // 3 pages: first two carry syncState pointing at the next page; third
    // omits syncState (end of history reached).
    const fetchFn = sequencedFetch([
      {
        messages: [
          { id: 'm9', sequenceId: 9 },
          { id: 'm8', sequenceId: 8 },
          { id: 'm7', sequenceId: 7 },
        ],
        _metadata: { syncState: syncUrl('cursor-2') },
      },
      {
        messages: [
          { id: 'm6', sequenceId: 6 },
          { id: 'm5', sequenceId: 5 },
          { id: 'm4', sequenceId: 4 },
        ],
        _metadata: { syncState: syncUrl('cursor-3') },
      },
      {
        messages: [
          { id: 'm3', sequenceId: 3 },
          { id: 'm2', sequenceId: 2 },
          { id: 'm1', sequenceId: 1 },
        ],
      }, // no _metadata.syncState
    ]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { chatId: 'c', pageSize: '3', maxPages: '10' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { messages: Array<{ id: string }>; hasMore: boolean; nextSyncState?: string; pagesFetched: number };
      expect(v.messages.map((m) => m.id)).toEqual(['m9', 'm8', 'm7', 'm6', 'm5', 'm4', 'm3', 'm2', 'm1']);
      expect(v.hasMore).toBe(false);
      expect(v.nextSyncState).toBeUndefined();
      expect(v.pagesFetched).toBe(3);
    }
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls).toHaveLength(3);
    // Second call should use the syncState cursor from page 1.
    expect(urls[1]).toContain('syncState=cursor-2');
    expect(urls[2]).toContain('syncState=cursor-3');
  });

  it('stops at --max-pages and sets hasMore:true (chains via nextSyncState) when the server keeps offering more', async () => {
    const fullPage = (cursor: string): Page => ({
      messages: [
        { id: 'a', sequenceId: 1 },
        { id: 'b', sequenceId: 2 },
      ],
      _metadata: { syncState: syncUrl(cursor) },
    });
    const fetchFn = sequencedFetch([fullPage('c2'), fullPage('c3'), fullPage('c4')]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { chatId: 'c', pageSize: '2', maxPages: '2' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { messages: unknown[]; hasMore: boolean; nextSyncState: string; pagesFetched: number };
      expect(v.messages).toHaveLength(4); // 2 pages × 2 messages
      expect(v.hasMore).toBe(true);
      expect(v.nextSyncState).toContain('syncState=c3');
      expect(v.pagesFetched).toBe(2);
    }
    expect((fetchFn as unknown as { urls: string[] }).urls).toHaveLength(2);
  });

  it('terminates on a zero-message page (server returned the syncState URL but it produced no rows)', async () => {
    // Edge case: server emits syncState but the next call comes back empty.
    // This can happen when the cursor points past the earliest message.
    const fetchFn = sequencedFetch([{ messages: [{ id: 'm1', sequenceId: 1 }], _metadata: { syncState: syncUrl('past-end') } }, { messages: [] }]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    const result = await cmd.execute(graph, { chatId: 'c', pageSize: '5', maxPages: '10' });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const v = result.value as { messages: unknown[]; hasMore: boolean; nextSyncState?: string };
      expect(v.messages).toHaveLength(1);
      expect(v.hasMore).toBe(false);
      expect(v.nextSyncState).toBeUndefined();
    }
    expect((fetchFn as unknown as { urls: string[] }).urls).toHaveLength(2);
  });

  it('uses --sync-state as the initial cursor (continues from a prior invocation)', async () => {
    const fetchFn = sequencedFetch([{ messages: [{ id: 'older', sequenceId: 50 }] }]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    await cmd.execute(graph, { chatId: 'c', syncState: syncUrl('continue-here') });
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls[0]).toContain('syncState=continue-here');
    // Should NOT include the default startTime=1 (we're using a server-provided cursor).
    // Actually the syncState URL DOES include startTime=1 because that's what the server bakes in.
    expect(urls[0]).toContain('startTime=1');
  });

  it('rejects --sync-state that is not a URL (validation_error before any fetch)', async () => {
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const r = await cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { chatId: 'c', syncState: 'not-a-url' });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.type).toBe('validation_error');
  });

  it('returns an error result when --sync-state URL does not match the expected substrate shape (server response format changed)', async () => {
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    // Valid URL but wrong host/path — should surface as a thrown error
    // (not a silent miscall against an unrelated endpoint).
    await expect(cmd.execute(createGraphClient(fakeAuth(), fakeFetch({})), { chatId: 'c', syncState: 'https://example.com/something' })).rejects.toThrow(
      'unexpected syncState URL shape'
    );
  });

  it('URI-encodes the chat id and routes via /api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages (first-call shape, no syncState)', async () => {
    const fetchFn = sequencedFetch([{ messages: [] }]);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fetchFn);
    await cmd.execute(graph, { chatId: '19:abc@thread.v2' });
    const urls = (fetchFn as unknown as { urls: string[] }).urls;
    expect(urls[0]).toContain('https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/19%3Aabc%40thread.v2/messages');
    expect(urls[0]).toContain('startTime=1');
    expect(urls[0]).toContain('pageSize=200');
    // The `|` separator in the view param is sent literally — matches what
    // Teams web emits (probed 2026-05-21 via Playwright bearer-trace).
    expect(urls[0]).toContain('view=msnp24Equivalent|supportsMessageProperties');
  });
});

// Audit Alex-session §A: a default `list-teams-chat-history` invocation used
// to return ~108 KB for 200 messages because every message carried the full
// IC3 envelope (annotations, threads, properties.policyViolation, etc.). The
// slim default projects each message to id/sequenceId/composetime/
// originalarrivaltime/messagetype/from/imdisplayname/content and truncates
// long content; callers that need the raw shape pass `--full true`.
describe('list-teams-chat-history applies slim projection by default', () => {
  // Use a fetch fake that returns the literal shape we care about — the
  // upstream IC3 substrate carries far more fields than the use-case keeps.
  const richMessageFetch = (messages: ReadonlyArray<Record<string, unknown>>): ((url: string) => Promise<Response>) => {
    let calls = 0;
    return async () => {
      calls += 1;
      const body = calls === 1 ? { messages } : { messages: [] };
      return new Response(JSON.stringify(body), { headers: { 'content-type': 'application/json' } });
    };
  };

  it('projects each message down to id/sequenceId/composetime/originalarrivaltime/messagetype/from/imdisplayname/content by default', async () => {
    const richMessage = {
      id: 'm1',
      sequenceId: 42,
      composetime: '2026-05-20T10:00:00Z',
      originalarrivaltime: '2026-05-20T10:00:00.001Z',
      messagetype: 'Text',
      from: '8:orgid:user-a',
      imdisplayname: 'Alice',
      content: 'hello world',
      // The fields below are present on every IC3 message but should NOT
      // appear in the slim projection.
      annotations: [{ kind: 'urgent' }],
      threadId: '19:abc@thread.v2',
      conversationLink: 'https://teams.microsoft.com/...',
      'properties.policyViolation': null,
    };
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), richMessageFetch([richMessage]));
    const result = await cmd.execute(graph, { chatId: '19:abc@thread.v2' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { messages: Array<Record<string, unknown>>; projection: string };
    expect(v.projection).toBe('slim');
    expect(v.messages).toHaveLength(1);
    expect(v.messages[0]).toEqual({
      id: 'm1',
      sequenceId: 42,
      composetime: '2026-05-20T10:00:00Z',
      originalarrivaltime: '2026-05-20T10:00:00.001Z',
      messagetype: 'Text',
      from: '8:orgid:user-a',
      imdisplayname: 'Alice',
      content: 'hello world',
    });
  });

  it('truncates content longer than --max-content-chars and marks the entry truncated:true with the original length', async () => {
    const longContent = 'a'.repeat(10_000);
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), richMessageFetch([{ id: 'm1', content: longContent }]));
    const result = await cmd.execute(graph, { chatId: 'c' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { messages: Array<Record<string, unknown>> };
    const msg = v.messages[0] ?? {};
    expect(typeof msg.content).toBe('string');
    expect((msg.content as string).length).toBe(4096);
    expect(msg.truncated).toBe(true);
    expect(msg.originalContentChars).toBe(10_000);
  });

  it('honors --max-content-chars override (cuts at the smaller cap)', async () => {
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), richMessageFetch([{ id: 'm1', content: 'abcdef' }]));
    const result = await cmd.execute(graph, { chatId: 'c', maxContentChars: '3' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { messages: Array<Record<string, unknown>> };
    const msg = v.messages[0] ?? {};
    expect(msg.content).toBe('abc');
    expect(msg.truncated).toBe(true);
    expect(msg.originalContentChars).toBe(6);
  });

  it('--full true returns the raw IC3 shape unchanged and reports projection:"full"', async () => {
    const richMessage = {
      id: 'm1',
      content: 'a'.repeat(10_000),
      annotations: [{ kind: 'urgent' }],
      threadId: '19:abc@thread.v2',
    };
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), richMessageFetch([richMessage]));
    const result = await cmd.execute(graph, { chatId: 'c', full: 'true' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as { messages: Array<Record<string, unknown>>; projection: string };
    expect(v.projection).toBe('full');
    expect(v.messages[0]).toEqual(richMessage);
  });

  it('rejects --full values other than true/false as a validation_error', async () => {
    const cmd = cmdMap['list-teams-chat-history'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch({ messages: [] }));
    const result = await cmd.execute(graph, { chatId: 'c', full: 'yes' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});

// Audit Alex-session §3 follow-up — get-excel-used-range slim projection.
// Raw Graph payload for usedRange() includes four parallel 2D arrays
// (values / text / numberFormat / formulas); the audit measured a 3×148
// sheet at 125 KB, mostly duplicate `"General"` numberFormat strings. The
// slim default ships `address` / `rowCount` / `columnCount` / `values` only;
// `--full true` opts back into the raw shape; `--max-cells` caps `values[]`
// size before the projection step.
describe('get-excel-used-range applies slim projection by default', () => {
  const richUsedRange = {
    address: 'Sheet1!A1:C3',
    rowCount: 3,
    columnCount: 3,
    values: [
      ['Q1', 'Q2', 'Q3'],
      [10, 20, 30],
      [40, 50, 60],
    ],
    text: [
      ['Q1', 'Q2', 'Q3'],
      ['10', '20', '30'],
      ['40', '50', '60'],
    ],
    // The audit's smoking gun — 9 cells of duplicated "General".
    numberFormat: [
      ['General', 'General', 'General'],
      ['General', 'General', 'General'],
      ['General', 'General', 'General'],
    ],
    formulas: [
      ['Q1', 'Q2', 'Q3'],
      [10, 20, 30],
      [40, 50, 60],
    ],
  };

  it("strips text/numberFormat/formulas by default, keeps address/rowCount/columnCount/values, reports projection:'slim'", async () => {
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch(richUsedRange));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as Record<string, unknown>;
    expect(v.projection).toBe('slim');
    expect(v.address).toBe('Sheet1!A1:C3');
    expect(v.rowCount).toBe(3);
    expect(v.columnCount).toBe(3);
    expect(v.values).toEqual(richUsedRange.values);
    expect(v.text).toBeUndefined();
    expect(v.numberFormat).toBeUndefined();
    expect(v.formulas).toBeUndefined();
  });

  it("`--full true` returns the raw Graph workbookRange shape (all four arrays) with projection:'full'", async () => {
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch(richUsedRange));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', full: 'true' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as Record<string, unknown>;
    expect(v.projection).toBe('full');
    expect(v.text).toEqual(richUsedRange.text);
    expect(v.numberFormat).toEqual(richUsedRange.numberFormat);
    expect(v.formulas).toEqual(richUsedRange.formulas);
  });

  it('drops `values[]` and surfaces a hint when usedRange exceeds --max-cells (regression guard against the 125 KB bloat case)', async () => {
    const giantSheet = { address: 'Sheet1!A1:Z2000', rowCount: 2000, columnCount: 26, values: Array.from({ length: 2000 }, () => Array.from({ length: 26 }, (_, i) => i)) };
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch(giantSheet));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', maxCells: '1000' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as Record<string, unknown>;
    expect(v.projection).toBe('slim');
    expect(v.truncated).toBe(true);
    expect(v.rowCount).toBe(2000);
    expect(v.columnCount).toBe(26);
    expect(v.values).toBeUndefined();
    expect((v.hint as string | undefined) ?? '').toContain('get-excel-range');
    expect((v.hint as string | undefined) ?? '').toContain('--full true');
  });

  it('--full true bypasses --max-cells (caller has opted into the full payload regardless of size)', async () => {
    const giantSheet = { address: 'Sheet1!A1:Z2000', rowCount: 2000, columnCount: 26, values: Array.from({ length: 2000 }, () => Array.from({ length: 26 }, () => 1)) };
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch(giantSheet));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', full: 'true', maxCells: '1' });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const v = result.value as Record<string, unknown>;
    expect(v.projection).toBe('full');
    expect(v.values).toBeDefined();
    expect(v.truncated).toBeUndefined();
  });

  it('rejects --full values other than true/false', async () => {
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch({}));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', full: 'yes' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('rejects non-positive --max-cells', async () => {
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    const graph = createGraphClient(fakeAuth(), fakeFetch({}));
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1', maxCells: '0' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('still translates WAC errors through mapWacError when the underlying item is not a workbook (excel-error.ts integration preserved)', async () => {
    const cmd = cmdMap['get-excel-used-range'];
    if (!cmd) throw new Error('command not found');
    // Return a 403 with the WAC needle in the body — mapWacError should rewrite it.
    const wacErrorFetch: FetchFn = async () =>
      new Response(JSON.stringify({ error: { code: 'AccessDenied', message: 'AccessDenied: Could not obtain a WAC access token.' } }), {
        status: 403,
        headers: { 'content-type': 'application/json' },
      });
    const graph = createGraphClient(fakeAuth(), wacErrorFetch);
    const result = await cmd.execute(graph, { driveId: 'd1', itemId: 'i1', worksheetId: 'Sheet1' });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('not an accessible Excel workbook');
    }
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
    const names = listChatMembers.meta.options.map((o) => o.name).toSorted((a, b) => a.localeCompare(b));
    expect(names).toEqual(['chat-id', 'filter', 'select', 'skip']);
  });

  it('list-chats.meta.options exposes top/skip/select/filter but NOT orderby/expand', () => {
    const names = listChats.meta.options.map((o) => o.name).toSorted((a, b) => a.localeCompare(b));
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
      // Audit Alex-session §2 follow-up: error is now the short DIAGNOSIS
      // ("--filter is incompatible..."); the actionable remedy (`list-mail-messages
      // --filter ...`) lives in the matching `hint` rule keyed by `code` so
      // the LLM gets remedy-in-hint, not diagnosis-and-remedy-in-error.
      expect(result.error.message).toContain('--filter');
      expect(result.error.message).toContain('SearchWithFilter');
      expect(result.error.message).not.toContain('list-mail-messages');
      if (result.error.type === 'validation_error') expect(result.error.code).toBe('cli_reject_search_with_filter');
    }
    expect(fetchFn.lastUrl).toBeNull();
  });

  it('still works when only --query is supplied (no regression on the happy path) — URL carries the slim default $select alongside $search', async () => {
    const url = await capturedUrl('search-mail-messages', { query: 'invoice' });
    expect(url).toBe(
      'https://graph.microsoft.com/v1.0/me/messages?$search="invoice"&$select=id%2Csubject%2Cfrom%2CtoRecipients%2CccRecipients%2CreceivedDateTime%2ChasAttachments%2CisRead%2Cimportance%2CbodyPreview'
    );
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
