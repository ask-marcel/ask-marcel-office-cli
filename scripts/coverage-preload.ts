/*
 * Coverage preload.
 *
 * `bun test --coverage` only reports rows for files the runner imports.
 * Untested infra, composition, and presenter files are silently absent
 * from the table, which makes the per-file gate trivially pass. This
 * preload side-effect-imports every such file so they appear at 0% (or
 * better) and the gate can fail loudly.
 *
 * Wired ONLY at coverage time, NOT in bunfig.toml. `scripts/check-coverage.ts`
 * spawns `bun test --coverage --preload ./scripts/coverage-preload.ts`. We do
 * NOT put `preload = [...]` under `[test]` in bunfig.toml because this file
 * pulls in heavy third-party SDKs (HTTP clients, cloud SDKs, AI clients,
 * loggers, etc. — whatever the infra adapters wrap) that would slow every
 * plain `bun test` by 1–2s.
 *
 * MAINTENANCE RULE: every new file under
 *   - src/infra/
 *   - src/composition/
 *   - src/presenter/
 * must be side-effect-imported here in the SAME commit that adds the file.
 * Reviewers check this explicitly. A pre-commit lint could enforce it; not
 * done yet, so it is a review obligation.
 *
 * See skills/atelier/references/workflow.md for the full rationale.
 */

import '../src/domain/jwt-utils.ts';
import '../src/infra/logger.ts';
import '../src/infra/auth.ts';
import '../src/infra/browser-auth.ts';
import '../src/infra/filesystem-bun.ts';
import '../src/infra/filesystem-node.ts';
import '../src/infra/graph-client.ts';
import '../src/infra/network-error.ts';
import '../src/infra/process-runner-bun.ts';
import '../src/infra/process-runner-node.ts';

import '../src/index.ts';
import '../src/composition/package-manager.ts';
import '../src/use-cases/commands/update.ts';

import '../src/composition/env.ts';
import '../src/composition/build-deps.ts';
import '../src/composition/cli.ts';

import '../src/presenter/output.ts';

import '../src/use-cases/commands/get-sharepoint-site.ts';
import '../src/use-cases/commands/list-sharepoint-site-drives.ts';
import '../src/use-cases/commands/get-sharepoint-site-drive-by-id.ts';
import '../src/use-cases/commands/list-sharepoint-site-lists.ts';
import '../src/use-cases/commands/get-sharepoint-site-list.ts';
import '../src/use-cases/commands/list-sharepoint-site-list-items.ts';
import '../src/use-cases/commands/get-sharepoint-site-list-item.ts';
import '../src/use-cases/commands/get-sharepoint-site-by-path.ts';
import '../src/use-cases/commands/list-todo-task-lists.ts';
import '../src/use-cases/commands/list-todo-tasks.ts';
import '../src/use-cases/commands/get-todo-task.ts';
import '../src/use-cases/commands/list-todo-linked-resources.ts';
import '../src/use-cases/commands/list-planner-tasks.ts';
import '../src/use-cases/commands/get-planner-plan.ts';
import '../src/use-cases/commands/list-plan-tasks.ts';
import '../src/use-cases/commands/get-planner-task.ts';
import '../src/use-cases/commands/get-planner-task-details.ts';
import '../src/use-cases/commands/list-plan-buckets.ts';
import '../src/use-cases/commands/get-planner-bucket.ts';
import '../src/use-cases/commands/list-mail-messages.ts';
import '../src/use-cases/commands/list-mail-folders.ts';
import '../src/use-cases/commands/list-mail-child-folders.ts';
import '../src/use-cases/commands/list-mail-folder-messages.ts';
import '../src/use-cases/commands/get-mail-message.ts';
import '../src/use-cases/commands/list-mail-attachments.ts';
import '../src/use-cases/commands/get-mail-attachment.ts';
import '../src/use-cases/commands/list-mail-rules.ts';
import '../src/use-cases/commands/get-mailbox-settings.ts';
import '../src/use-cases/commands/list-onenote-notebooks.ts';
import '../src/use-cases/commands/list-onenote-notebook-sections.ts';
import '../src/use-cases/commands/list-all-onenote-sections.ts';
import '../src/use-cases/commands/list-onenote-section-pages.ts';
import '../src/use-cases/commands/get-onenote-page-as-markdown.ts';
import '../src/use-cases/commands/get-onenote-page-content.ts';
import '../src/use-cases/commands/download-drive-item-as-markdown.ts';
import '../src/use-cases/commands/extract-drive-item-images.ts';
import '../src/use-cases/commands/extract-mail-attachment-images.ts';
import '../src/use-cases/commands/media-files.ts';
import '../src/use-cases/commands/download-drive-item-as-pdf.ts';
import '../src/use-cases/commands/download-drive-item-version.ts';
import '../src/infra/turndown-adapter.ts';
import '../src/infra/ooxml-zip-adapter.ts';
import '../src/infra/ooxml-media-extractor.ts';
import '../src/use-cases/commands/inline-image-embedder.ts';
import '../src/use-cases/commands/markdown-pipeline.ts';
import '../src/use-cases/commands/text-passthrough.ts';
import '../src/use-cases/commands/sharepoint-link-extractor.ts';
import '../src/use-cases/commands/extract-sharepoint-links-in-mail.ts';
import '../src/use-cases/commands/convert-mail-to-markdown.ts';
import '../src/use-cases/commands/convert-mail-attachment-to-pdf.ts';
import '../src/use-cases/commands/convert-mail-attachment-to-markdown.ts';
import '../src/use-cases/commands/embedded-item-to-markdown.ts';
import '../src/use-cases/commands/get-current-user.ts';
import '../src/use-cases/commands/get-my-profile-photo.ts';
import '../src/use-cases/commands/list-calendar-events.ts';
import '../src/use-cases/commands/get-calendar-event.ts';
import '../src/use-cases/commands/list-specific-calendar-events.ts';
import '../src/use-cases/commands/get-specific-calendar-event.ts';
import '../src/use-cases/commands/get-calendar-view.ts';
import '../src/use-cases/commands/get-specific-calendar-view.ts';
import '../src/use-cases/commands/list-calendar-event-instances.ts';
import '../src/use-cases/commands/list-calendars.ts';
import '../src/use-cases/commands/list-calendar-events-delta.ts';
import '../src/use-cases/commands/list-calendar-view-delta.ts';
import '../src/use-cases/commands/list-chat-members.ts';
import '../src/use-cases/commands/list-joined-teams.ts';
import '../src/use-cases/commands/get-team.ts';
import '../src/use-cases/commands/list-team-channels.ts';
import '../src/use-cases/commands/get-team-channel.ts';
import '../src/use-cases/commands/list-teams-chats-with-messages.ts';
import '../src/use-cases/commands/list-teams-chat-messages.ts';
import '../src/use-cases/commands/get-teams-chat-message.ts';
