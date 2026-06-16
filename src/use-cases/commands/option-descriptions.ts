// Shared option descriptions. When the same option (here `--drive-id`) recurs
// across many commands, keep ONE canonical wording so the discovery guidance
// can't drift per-command (QA-003 / friction-plan #5). Commands with genuinely
// command-specific context (Excel "containing the workbook", search "to search
// inside", the site-scoped get-sharepoint-site-drive-by-id) keep their tailored
// descriptions on purpose — only the generic drive-item commands share this.

/**
 * `--drive-id` description for every generic OneDrive / SharePoint drive-item
 * command. Always points the caller at how to obtain a drive id.
 */
export const DRIVE_ID_DESCRIPTION =
  'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.';
