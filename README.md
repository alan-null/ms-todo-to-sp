# MS Todo to Super Productivity Migration Tools

This repository contains PowerShell scripts for migrating data from Microsoft To Do to **Super Productivity** (SP), and for merging existing SP backups.

## Scripts Overview

| Script                                           | Description                                                                   |
| ------------------------------------------------ | ----------------------------------------------------------------------------- |
| [Convert-MsTodoToSP.ps1](Convert-MsTodoToSP.ps1) | Converts Microsoft To Do JSON exports to SP compatible format.                |
| [Merge-SPBackups.ps1](Merge-SPBackups.ps1)       | Merges two Super Productivity backup JSON files into a single unified backup. |

## Table of Contents

- [1. Microsoft To Do to Super Productivity Migration](#1-microsoft-to-do-to-super-productivity-migration)
  - [Overview](#overview)
  - [Usage](#usage)
  - [Object Mappings](#object-mappings)
    - [Projects (MS Todo Lists → SP Projects)](#projects-ms-todo-lists--sp-projects)
    - [Tasks (MS Todo Tasks → SP Tasks)](#tasks-ms-todo-tasks--sp-tasks)
    - [Subtasks (MS Todo checklistItems → SP Subtasks)](#subtasks-ms-todo-checklistitems--sp-subtasks)
    - [Recurrence Configurations (MS Todo recurrence → SP taskRepeatCfg)](#recurrence-configurations-ms-todo-recurrence--sp-taskrepeatcfg)
  - [Example Output](#example-output)
- [2. Super Productivity Backup Merger](#2-super-productivity-backup-merger)
  - [Overview](#overview-1)
  - [Usage](#usage-1)
  - [Parameters](#parameters)
  - [Duplicate Detection](#duplicate-detection)
    - [Projects](#projects)
    - [Tags](#tags)
    - [Tasks](#tasks)
  - [Example Output](#example-output-1)


---

## 1. Microsoft To Do to Super Productivity Migration

### Overview
The [Convert-MsTodoToSP.ps1](Convert-MsTodoToSP.ps1) script handles comprehensive data migration including `tasks`, `projects`, `subtasks`, `tags`, `due dates`, `reminders`, and `recurrence` patterns.

> All IDs are 21-character base64url strings (nanoid-style), not GUIDs.

### Usage
```powershell
.\Convert-MsTodoToSP.ps1 -InputFile "path\to\ms-todo.json" -OutputFile "output.json"
```

### Object Mappings

#### Projects (MS Todo Lists → SP Projects)
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

#### Tasks (MS Todo Tasks → SP Tasks)
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

#### Subtasks (MS Todo checklistItems → SP Subtasks)
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

#### Recurrence Configurations (MS Todo recurrence → SP taskRepeatCfg)
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


### Example Output

```powershell
pwsh .\Convert-MsTodoToSP.ps1 -InputFile .\input\microsoft-todo-backup.json -OutputFile ".\out\ms-converted.json"

Conversion complete.
  Projects : 29
  Tasks    : 2005 (incl. subtasks)
  Tags     : 13 (+ TODAY tag)
  Repeats  : 39
  Reminders: 24
  Output   : C:\repo\ms-todo-to-sp\out\ms-converted.json
```

---

## 2. Super Productivity Backup Merger

### Overview
The [Merge-SPBackups.ps1](Merge-SPBackups.ps1) script merges two **Super Productivity** (SP) backup JSON files into a single unified backup.

The **Primary** file is the base — its data is preserved — and unique data from the **Secondary** file is merged in with full deduplication and conflict resolution.

Designed for SP backup format version **4.5** (`crossModelVersion`).

### Usage
```powershell
.\Merge-SPBackups.ps1 -PrimaryFile "a.json" -SecondaryFile "b.json" -OutputFile "merged.json"
```

### Parameters

| Parameter        | Default                                | Description                                |
| ---------------- | -------------------------------------- | ------------------------------------------ |
| `-PrimaryFile`   | `input\super-productivity-backup.json` | Base backup file (its data takes priority) |
| `-SecondaryFile` | `out\dist.json`                        | Backup to merge into the primary           |
| `-OutputFile`    | `out\merged.json`                      | Path for the merged output file            |

### Duplicate Detection

Every entity type uses a two-tier matching strategy:

#### Projects
1. **By ID** — exact match of project ID
2. **By title** — case-insensitive title comparison

#### Tags
1. **By ID** — exact match (including well-known `"TODAY"`)
2. **By title** — case-insensitive title comparison

#### Tasks
1. **By ID** — exact match of task ID
2. **By title + project** — case-insensitive title within the same (mapped) project
3. **Instance guard** — title matches are confirmed via `created`, `doneOn`, `dueDay`, etc.

### Example Output

```powershell
pwsh .\Merge-SPBackups.ps1 -PrimaryFile .\input\sp-backup.json -SecondaryFile .\out\ms-converted.json -OutputFile .\out\sp-and-ms.json

=============================================
 Merge complete
=============================================

  Source stats:
                  Primary  Secondary
    Projects:           3         29
    Tasks:              2       2005
    Subtasks:           0          0
    Tags:               1         13
    Repeats:            0         39
    Reminders:          0         24

  Merge actions:
    Projects  : 0 matched, 29 added
    Tags      : 0 matched, 13 added
    Tasks     : 0 merged, 1317 added, 0 unchanged
    Subtasks  : 0 merged, 688 added
    RepeatCfgs: 0 matched, 39 added
    Reminders : 24 added

  Output totals:
    Projects  : 32
    Tasks     : 2007 (incl. subtasks)
    Tags      : 14 (+ TODAY)
    RepeatCfgs: 39
    Reminders : 24
    Output    : C:\repo\ms-todo-to-sp\out\sp-and-ms.json
```