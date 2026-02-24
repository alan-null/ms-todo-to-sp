# MS Todo to Super Productivity Migration

## Overview
This document summarizes the PowerShell script developed to convert **Microsoft To Do** JSON exports to **Super Productivity** compatible format.

The script handles comprehensive data migration including `tasks`, `projects`, `subtasks`, `tags`, `due dates`, `reminders`, and `recurrence` patterns.

> All IDs are 21-character base64url strings (nanoid-style), not GUIDs.

## Object Mappings

### Projects (MS Todo Lists → SP Projects)
MS Todo lists are converted to SP projects and grouped under a single `"MS-Imported"` folder in the menu tree.

| MS Todo List Field | SP Project Field   | Notes                                              |
| ------------------ | ------------------ | -------------------------------------------------- |
| `displayName`      | `title`            | List name becomes project title                    |
| N/A                | `id`               | Random 21-char base64url ID                        |
| N/A                | `taskIds`          | Array of task IDs belonging to this project        |
| N/A                | `advancedCfg`      | Worklog export settings (same as default projects) |
| N/A                | `theme`            | Default project theme (`#29a1aa`)                  |
| N/A                | `icon`             | `"list_alt"`                                       |
| N/A                | `isHiddenFromMenu` | `false`                                            |
| N/A                | `isArchived`       | `false`                                            |
| N/A                | `isEnableBacklog`  | `false`                                            |
| N/A                | `backlogTaskIds`   | Empty array                                        |
| N/A                | `noteIds`          | Empty array                                        |

### Tasks (MS Todo Tasks → SP Tasks)
Core task data with all MS Todo features mapped.

| MS Todo Task Field                           | SP Task Field    | Notes                                                                                                |
| -------------------------------------------- | ---------------- | ---------------------------------------------------------------------------------------------------- |
| `title`                                      | `title`          | Direct mapping; inline `#tags` are also extracted (title is kept as-is)                              |
| `status`                                     | `isDone`         | `"completed"` → `true`, all other values → `false`                                                   |
| `createdDateTime`                            | `created`        | Converted to Unix timestamp (ms); falls back to script run time                                      |
| `lastModifiedDateTime`                       | `modified`       | Converted to Unix timestamp (ms); falls back to script run time                                      |
| `completedDateTime.dateTime`                 | `doneOn`         | Unix ms timestamp; falls back to `modified` when task is done but field is absent                    |
| `dueDateTime.dateTime`                       | `dueDay`         | YYYY-MM-DD string — set when time-of-day component is < 60 seconds (i.e. date-only)                  |
| `dueDateTime.dateTime`                       | `dueWithTime`    | Unix ms timestamp — set when time-of-day component is ≥ 60 seconds; mutually exclusive with `dueDay` |
| `isReminderOn` + `reminderDateTime.dateTime` | `remindAt`       | Absolute Unix ms timestamp stored directly; `reminderId` is also set                                 |
| `recurrence`                                 | `repeatCfgId`    | ID of the created `taskRepeatCfg` entity; only set for incomplete tasks                              |
| `categories`                                 | `tagIds`         | Each category maps to a tag ID                                                                       |
| `importance = "high"`                        | `tagIds`         | Adds the `"Important"` tag ID                                                                        |
| `#word` in `title`                           | `tagIds`         | Inline hashtags in the title are linked as tags                                                      |
| `body.content`                               | `notes`          | Trimmed text content; omitted if blank                                                               |
| N/A                                          | `id`             | Random 21-char base64url ID                                                                          |
| N/A                                          | `projectId`      | References parent project                                                                            |
| N/A                                          | `subTaskIds`     | Array of subtask IDs                                                                                 |
| N/A                                          | `timeSpentOnDay` | Empty object                                                                                         |
| N/A                                          | `timeEstimate`   | `0`                                                                                                  |
| N/A                                          | `timeSpent`      | `0`                                                                                                  |
| N/A                                          | `hasPlannedTime` | `true` if `dueDay` or `dueWithTime` is set, otherwise `false`                                        |
| N/A                                          | `attachments`    | Empty array                                                                                          |

### Subtasks (MS Todo checklistItems → SP Subtasks)
Checklist items become subtasks linked to their parent task.

| MS Todo checklistItem Field | SP Subtask Field | Notes                                     |
| --------------------------- | ---------------- | ----------------------------------------- |
| `displayName`               | `title`          | Checklist item text                       |
| `isChecked`                 | `isDone`         | Direct boolean mapping                    |
| `createdDateTime`           | `created`        | Unix ms timestamp; falls back to `$NowMs` |
| `createdDateTime`           | `modified`       | Same value as `created`                   |
| N/A                         | `id`             | Random 21-char base64url ID               |
| N/A                         | `projectId`      | Same as parent task                       |
| N/A                         | `parentId`       | Parent task ID                            |
| N/A                         | `subTaskIds`     | Empty array                               |
| N/A                         | `tagIds`         | Empty array                               |
| N/A                         | `timeSpentOnDay` | Empty object                              |
| N/A                         | `timeEstimate`   | `0`                                       |
| N/A                         | `timeSpent`      | `0`                                       |
| N/A                         | `hasPlannedTime` | `false`                                   |
| N/A                         | `attachments`    | Empty array                               |

### Tags (MS Todo Sources → SP Tags)
Tags are collected from three sources in a first-pass scan: `categories`, `importance = "high"` (→ `"Important"`), and inline `#tags` in task titles. A mandatory `TODAY` tag (id = `"TODAY"`) is always created.

| Source                  | SP Tag Field  | Notes                                               |
| ----------------------- | ------------- | --------------------------------------------------- |
| Category name / hashtag | `title`       | Direct mapping                                      |
| N/A                     | `id`          | Random 21-char base64url ID (assigned at runtime)   |
| N/A                     | `taskIds`     | Back-filled with IDs of tasks that use the tag      |
| N/A                     | `color`       | `null`                                              |
| N/A                     | `created`     | Script run timestamp                                |
| N/A                     | `modified`    | Script run timestamp                                |
| N/A                     | `icon`        | `null` (imported tags); `"wb_sunny"` for TODAY      |
| N/A                     | `advancedCfg` | Default worklog export settings                     |
| N/A                     | `theme`       | Default tag theme (`#a05db1`); TODAY uses `#6495ED` |

### Recurrence Configurations (MS Todo recurrence → SP taskRepeatCfg)
Repeat configs are created only for **incomplete** tasks that have a `recurrence.pattern`. Completed recurring tasks are treated as historical instances and migrated as plain done tasks.

| MS Todo Recurrence Field                       | SP RepeatCfg Field         | Notes                                                                                                                                                                                           |
| ---------------------------------------------- | -------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `pattern.type`                                 | `repeatCycle`              | `"daily"` → `DAILY`, `"weekly"` → `WEEKLY`, `"absoluteMonthly"`/`"relativeMonthly"` → `MONTHLY`, `"absoluteYearly"`/`"relativeYearly"` → `YEARLY`, `"hourly"` → `DAILY` (with warning)          |
| `pattern.interval`                             | `repeatEvery`              | Direct mapping; clamped to minimum 1                                                                                                                                                            |
| `pattern.daysOfWeek`                           | `monday`…`sunday`          | Boolean flags; falls back to the weekday of `createdDate` if absent (weekly only)                                                                                                               |
| dueDate / `pattern.dayOfMonth` / `createdDate` | `startDate`                | MONTHLY/YEARLY only. Priority: dueDate > `dayOfMonth` field > `createdDate`                                                                                                                     |
| task `title`                                   | `title`                    | Copied from the source task                                                                                                                                                                     |
| task `notes`                                   | `notes`                    | Copied from the source task                                                                                                                                                                     |
| task `tagIds`                                  | `tagIds`                   | Snapshot of the task's tag IDs at conversion time                                                                                                                                               |
| N/A                                            | `id`                       | Random 21-char base64url ID                                                                                                                                                                     |
| N/A                                            | `projectId`                | Same as the source task's project                                                                                                                                                               |
| N/A                                            | `lastTaskCreation`         | Source task's `created` timestamp                                                                                                                                                               |
| N/A                                            | `lastTaskCreationDay`      | YYYY-MM-DD of `lastTaskCreation`                                                                                                                                                                |
| N/A                                            | `order`                    | Global incrementing counter (insertion order)                                                                                                                                                   |
| N/A                                            | `quickSetting`             | `DAILY` (daily×1), `WEEKLY_CURRENT_WEEKDAY` (weekly×1, 1 day), `MONDAY_TO_FRIDAY` (weekly×1, Mon–Fri), `MONTHLY_CURRENT_DATE` (monthly×1), `YEARLY_CURRENT_DATE` (yearly×1), otherwise `CUSTOM` |
| N/A                                            | `isPaused`                 | `false`                                                                                                                                                                                         |
| N/A                                            | `shouldInheritSubtasks`    | `false`                                                                                                                                                                                         |
| N/A                                            | `repeatFromCompletionDate` | `false`                                                                                                                                                                                         |
| N/A                                            | `deletedInstanceDates`     | Empty array                                                                                                                                                                                     |

### Reminders
A SP reminder entity is created when `isReminderOn = true` and `reminderDateTime.dateTime` is present and parseable. The absolute time is stored directly — no relative-offset mapping is applied.

| MS Todo Field               | SP Reminder Field | Notes                       |
| --------------------------- | ----------------- | --------------------------- |
| `reminderDateTime.dateTime` | `remindAt`        | Absolute Unix ms timestamp  |
| `title`                     | `title`           | Copied from the task title  |
| N/A                         | `id`              | Random 21-char base64url ID |
| N/A                         | `type`            | `"TASK"`                    |
| N/A                         | `relatedId`       | The task's ID               |

## Organizational Structure
- **Project folder**: All imported projects are placed under a single `"MS-Imported"` folder in the project menu tree.
- **Tag folder**: Imported tags (excluding `TODAY`) are placed under a `"MS-Imported Tags"` folder in the tag menu tree. The `TODAY` tag always appears at the root of the tag tree.
- **Tag ID ordering**: Imported tags are stored alphabetically (case-insensitive); `TODAY` is appended last.

## Data Integrity Features
- Graceful timestamp parsing — falls back to `$NowMs` on invalid or missing values
- Complete entity structures to prevent SP UI errors on import
- Bidirectional task↔tag relationships (`task.tagIds` ↔ `tag.taskIds`)
- `List<string>` collections are converted to plain arrays before JSON serialisation to avoid `{}` objects in output
- UTF-8 encoding for international characters

## Usage
```powershell
.\Convert-MsTodoToSP.ps1 -InputFile "path\to\ms-todo.json" -OutputFile "output.json"
```