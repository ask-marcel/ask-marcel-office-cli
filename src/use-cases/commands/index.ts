import type { Command } from './command-types.ts';
import * as downloadDriveItemAsMarkdown from './download-drive-item-as-markdown.ts';
import * as extractDriveItemImages from './extract-drive-item-images.ts';
import * as listAccessibleDrives from './list-accessible-drives.ts';
import * as downloadDriveItemAsPdf from './download-drive-item-as-pdf.ts';
import * as downloadDriveItemVersion from './download-drive-item-version.ts';
import * as downloadOnedriveFileContent from './download-onedrive-file-content.ts';
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
import * as convertMailAttachmentToMarkdown from './convert-mail-attachment-to-markdown.ts';
import * as extractMailAttachmentImages from './extract-mail-attachment-images.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as listCalendarEventAttachments from './list-calendar-event-attachments.ts';
import * as convertCalendarEventAttachmentToMarkdown from './convert-calendar-event-attachment-to-markdown.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';
import * as convertMailToMarkdown from './convert-mail-to-markdown.ts';
import * as convertDriveItemZip from './convert-drive-item-zip.ts';
import * as convertMailAttachmentZip from './convert-mail-attachment-zip.ts';
import * as extractSharepointLinksInDocuments from './extract-sharepoint-links-in-documents.ts';
import * as extractSharepointLinksInMail from './extract-sharepoint-links-in-mail.ts';
import * as listChats from './list-chats.ts';
import * as getChat from './get-chat.ts';
import * as listTeamsChatsWithMessages from './list-teams-chats-with-messages.ts';
import * as listTeamsChatMessages from './list-teams-chat-messages.ts';
import * as listTeamsChatHistory from './list-teams-chat-history.ts';
import * as getTeamsChatMessage from './get-teams-chat-message.ts';
import * as resolveTeamsLink from './resolve-teams-link.ts';
import * as resolveMailLink from './resolve-mail-link.ts';
import * as resolveDriveShareLink from './resolve-drive-share-link.ts';
import * as resolveCalendarLink from './resolve-calendar-link.ts';
import * as findChatsWithUser from './find-chats-with-user.ts';
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
import * as myQuickContext from './my-quick-context.ts';
import * as scopesCheck from './scopes-check.ts';
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

const commands: Record<string, Command> = {
  'list-drives': listDrives,
  'get-drive-root-item': getDriveRootItem,
  'list-folder-files': listFolderFiles,
  'download-onedrive-file-content': downloadOnedriveFileContent,
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
  'extract-sharepoint-links-in-documents': extractSharepointLinksInDocuments,
  'convert-mail-to-markdown': convertMailToMarkdown,
  'convert-drive-item-zip': convertDriveItemZip,
  'convert-mail-attachment-zip': convertMailAttachmentZip,
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
  'my-quick-context': myQuickContext,
  'scopes-check': scopesCheck,
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

export { commands };
