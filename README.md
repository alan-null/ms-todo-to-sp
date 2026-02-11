# MS Todo to Super Productivity Migration

## Overview
This document summarizes the PowerShell script developed to convert **Microsoft To Do** JSON exports to **Super Productivity** compatible format.

The script handles comprehensive data migration including `tasks`, `projects`, `subtasks`, `tags`, `due dates`, `reminders`, and `recurrence` patterns.


## Object Mappings

### Projects (MS Todo Lists → SP Projects)
MS Todo lists are converted to SP projects with organized folder structure.

| MS Todo List Field | SP Project Field   | Notes                                              |
| ------------------ | ------------------ | -------------------------------------------------- |
| `displayName`      | `title`            | List name becomes project title                    |
| N/A                | `id`               | Generated as random GUID                           |
| N/A                | `taskIds`          | Array of task IDs belonging to this project        |
| N/A                | `advancedCfg`      | Worklog export settings (same as default projects) |
| N/A                | `theme`            | Default theme configuration                        |
| N/A                | `icon`             | Set to "list"                                      |
| N/A                | `isHiddenFromMenu` | `false`                                              |
| N/A                | `isArchived`       | `false`                                              |
| N/A                | `isEnableBacklog`  | `false`                                              |
| N/A                | `backlogTaskIds`   | Empty array                                        |
| N/A                | `noteIds`          | Empty array                                        |

### Tasks (MS Todo Tasks → SP Tasks)
Core task data with all MS Todo features mapped.

| MS Todo Task Field     | SP Task Field    | Notes                                                  |
| ---------------------- | ---------------- | ------------------------------------------------------ |
| `title`                | `title`          | Direct mapping                                         |
| `status`               | `isDone`         | "completed" → `true`, others → `false`                     |
| `createdDateTime`      | `created`        | Converted to Unix timestamp (ms)                       |
| `lastModifiedDateTime` | `modified`       | Converted to Unix timestamp (ms)                       |
| `completedDateTime`    | `doneOn`         | Converted to Unix timestamp (ms) if completed          |
| `dueDateTime.dateTime` | `dueDay`         | Date part in YYYY-MM-DD format                         |
| `dueDateTime.dateTime` | `dueWithTime`    | Boolean: `true` if time specified (not 00:00:00)         |
| `reminderDateTime`     | `remindAt`       | Mapped to SP reminder options (AtStart, m5, m10, etc.) |
| `recurrence`           | `repeatCfgId`    | References created repeat configuration                |
| `categories`           | `tagIds`         | Array of tag IDs for categories                        |
| `importance`           | `tagIds`         | "high" adds "Important" tag                            |
| N/A                    | `id`             | Generated as random GUID                               |
| N/A                    | `projectId`      | References parent project                              |
| N/A                    | `subTaskIds`     | Array of subtask IDs                                   |
| N/A                    | `timeSpentOnDay` | Empty object                                           |
| N/A                    | `timeEstimate`   | 0                                                      |
| N/A                    | `timeSpent`      | 0                                                      |
| N/A                    | `hasPlannedTime` | `false`                                                  |
| N/A                    | `attachments`    | Empty array                                            |
| N/A                    | `parentId`       | null (for main tasks)                                  |
| N/A                    | `reminderId`     | null                                                   |
| N/A                    | `repeatCfgId`    | Set if recurrence exists                               |

### Subtasks (MS Todo checklistItems → SP Subtasks)
Checklist items become subtasks.

| MS Todo checklistItem Field | SP Subtask Field | Notes                       |
| --------------------------- | ---------------- | --------------------------- |
| `displayName`               | `title`          | Checklist item text         |
| `isChecked`                 | `isDone`         | Direct boolean mapping      |
| `createdDateTime`           | `created`        | Converted to Unix timestamp |
| `createdDateTime`           | `modified`       | Same as created             |
| N/A                         | `id`             | Generated as random GUID    |
| N/A                         | `projectId`      | Same as parent task         |
| N/A                         | `parentId`       | Parent task ID              |
| N/A                         | `subTaskIds`     | Empty array                 |
| N/A                         | `tagIds`         | Empty array                 |
| N/A                         | `timeSpentOnDay` | Empty object                |
| N/A                         | `timeEstimate`   | 0                           |
| N/A                         | `timeSpent`      | 0                           |
| N/A                         | `hasPlannedTime` | `false`                       |
| N/A                         | `attachments`    | Empty array                 |

### Tags (MS Todo Categories → SP Tags)
Categories become tags with deterministic IDs.

| MS Todo Category       | SP Tag Field  | Notes                            |
| ---------------------- | ------------- | -------------------------------- |
| Category name (string) | `title`       | Direct mapping                   |
| N/A                    | `id`          | Deterministic GUID based on name |
| N/A                    | `taskIds`     | Array of task IDs using this tag |
| N/A                    | `color`       | null                             |
| N/A                    | `created`     | Current timestamp                |
| N/A                    | `modified`    | Current timestamp                |
| N/A                    | `icon`        | null                             |
| N/A                    | `advancedCfg` | Worklog export settings          |
| N/A                    | `theme`       | Default theme configuration      |

### Recurrence Configurations (MS Todo recurrence → SP taskRepeatCfg)
Recurring tasks create separate repeat configuration entities.

| MS Todo Recurrence Field | SP RepeatCfg Field         | Notes                                                  |
| ------------------------ | -------------------------- | ------------------------------------------------------ |
| `pattern.type`           | `repeatCycle`              | "daily"→"D", "weekly"→"W", "monthly"→"M", "yearly"→"Y" |
| `pattern.interval`       | `repeatEvery`              | Direct mapping                                         |
| `pattern.daysOfWeek`     | `monday`/`tuesday`/etc.    | Boolean flags for each day                             |
| `range.startDate`        | `startDate`                | Date part of start                                     |
| `range.endDate`          | `endDate`                  | Date if not "noEnd"                                    |
| N/A                      | `id`                       | Generated as random GUID                               |
| N/A                      | `title`                    | "Imported from MS Todo"                                |
| N/A                      | `startTime`                | null                                                   |
| N/A                      | `lastTaskCreation`         | Current timestamp                                      |
| N/A                      | `lastTaskCreationDay`      | Current date                                           |
| N/A                      | `isPaused`                 | `false`                                                  |
| N/A                      | `shouldInheritSubtasks`    | `false`                                                  |
| N/A                      | `repeatFromCompletionDate` | `false`                                                  |
| N/A                      | `quickSetting`             | "WEEKLY_CURRENT_WEEKDAY" for weekly                    |

## Reminder Mapping Logic
MS Todo absolute reminder times are mapped to SP's relative reminder options:

- If reminder time equals due time: `AtStart`
- If reminder is 5-10 minutes before due: `m5`, `m10`
- If reminder is 15-60 minutes before due: `m15`, `m30`, `h1`
- If no due date, compares to creation time for `AtStart`
- Otherwise: `DoNotRemind`

## Organizational Structure
- **Project Folders**: All imported projects placed in "MS-IMPORT-PROJECTS" menu folder
- **Tag Folders**: All imported tags placed in "MS-IMPORT-TAGS" menu folder
- **Deterministic IDs**: Tags use consistent IDs based on names to prevent re-import conflicts

## Data Integrity Features
- Handles invalid dates in recurrence ranges
- Graceful parsing of timestamps
- Complete entity structures to prevent UI errors
- Bidirectional task-tag relationships
- UTF-8 encoding for international characters

## Usage
```powershell
.\Convert-MsTodoToSP.ps1 -InputFile "path\to\ms-todo.json" -OutputFile "output.json"
```