# PowerShell script to convert Microsoft To Do JSON export to Super Productivity format
# Usage: .\Convert-MsTodoToSP.ps1 -InputFile "temp\ms-todo.json" -OutputFile "temp\sp_converted.json"

param(
    [string]$InputFile = "temp\ms-todo.json",
    [string]$OutputFile = "temp\sp_converted.json"
)

# Function to generate a short unique ID similar to SP format
function New-SPId {
    # Generate a 21-character ID using base64url encoding of a GUID
    $guid = [guid]::NewGuid()
    $bytes = $guid.ToByteArray()
    $base64 = [Convert]::ToBase64String($bytes)
    # Make it URL-safe and take first 21 chars
    $base64 = $base64.Replace('+', '-').Replace('/', '_').Replace('=', '')
    return $base64.Substring(0, 21)
}

# Function to get DB date string (YYYY-MM-DD)
function Get-DbDateStr {
    param([DateTimeOffset]$date)
    return $date.ToString("yyyy-MM-dd")
}

function Add-IfNotNull {
    param($hashtable, $key, $value)
    if ($null -ne $value) {
        $hashtable[$key] = $value
    }
}

$msTodo = Get-Content -Path $InputFile -Encoding UTF8 | ConvertFrom-Json

# Collect imported project and tag ids
$importedProjectIds = @()

# Collect unique tags
$tagNames = @{}
$tagNames["Important"] = $true  # For high importance tasks

# Process lists and tasks to collect categories
foreach ($list in $msTodo.lists) {
    if ($msTodo.tasks.PSObject.Properties[$list.id]) {
        foreach ($msTask in $msTodo.tasks.($list.id)) {
            if ($msTask.categories) {
                foreach ($cat in $msTask.categories) {
                    $tagNames[$cat] = $true
                }
            }
        }
    }
}

# Create tag entities
$tagEntities = @{}
foreach ($tagName in $tagNames.Keys) {
    $tagId = New-SPId
    $tagEntities[$tagId] = @{
        id          = $tagId
        title       = $tagName
        color       = "#42a5f5"  # Default color # TODO: Hardcoded Default Values
        created     = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()
        modified    = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()
        icon        = $null
        taskIds     = @()
        advancedCfg = @{
            worklogExportSettings = @{
                cols             = @("DATE", "START", "END", "TIME_CLOCK", "TITLES_INCLUDING_SUB")
                roundWorkTimeTo  = $null
                roundStartTimeTo = $null
                roundEndTimeTo   = $null
                separateTasksBy  = " | "
                groupBy          = "DATE"
            }
        }
        theme       = @{
            isAutoContrast           = $true
            isDisableBackgroundTint  = $false
            primary                  = "#42a5f5"
            huePrimary               = "500"
            accent                   = "#ff4081"
            hueAccent                = "500"
            warn                     = "#e11826"
            hueWarn                  = "500"
            backgroundImageDark      = $null
            backgroundImageLight     = $null
            backgroundOverlayOpacity = 0
        }
    }
}

# Set imported tag ids
$importedTagIds = $tagEntities.Keys

# Initialize SP data structure
$sp = @{
    task          = @{
        ids                   = @()
        entities              = @{}
        currentTaskId         = $null
        selectedTaskId        = $null
        taskDetailTargetPanel = $null
        lastCurrentTaskId     = $null
        isDataLoaded          = $true
    }
    project       = @{
        ids      = @()
        entities = @{}
    }
    tag           = @{
        ids      = $tagEntities.Keys
        entities = $tagEntities
    }
    taskRepeatCfg = @{
        ids      = @()
        entities = @{}
    }
    simpleCounter = @{
        ids      = @()
        entities = @{}
    }
    metric        = @{
        ids      = @()
        entities = @{}
    }
    reminders     = @()
    planner       = @{
        ids      = @()
        entities = @{}
        days     = @{}
    }
    boards        = @{
        ids       = @()
        entities  = @{}
        boardCfgs = @()
    }
    note          = @{
        ids        = @()
        entities   = @{}
        todayOrder = @()
    }
    issueProvider = @{
        ids      = @()
        entities = @{}
    }
    menuTree      = @{
        projectTree = @()
        tagTree     = @()
    }
    globalConfig  = @{
        appFeatures  = @{
            isTimeTrackingEnabled     = $true
            isFocusModeEnabled        = $true
            isSchedulerEnabled        = $true
            isPlannerEnabled          = $true
            isBoardsEnabled           = $true
            isScheduleDayPanelEnabled = $true
            isIssuesPanelEnabled      = $true
            isProjectNotesEnabled     = $true
            isSyncIconEnabled         = $true
            isDonatePageEnabled       = $true
            isEnableUserProfiles      = $true
            isHabitsEnabled           = $true
        }
        localization = @{}
        misc         = @{
            isConfirmBeforeExit                 = $false
            isConfirmBeforeExitWithoutFinishDay = $false
            isMinimizeToTray                    = $false
            startOfNextDay                      = 0
            isDisableAnimations                 = $false
        }
        tasks        = @{
            isAutoMarkParentAsDone             = $false
            isAutoAddWorkedOnToToday           = $false
            isTrayShowCurrent                  = $false
            isMarkdownFormattingInNotesEnabled = $true
            notesTemplate                      = ""
        }
        shortSyntax  = @{
            isEnableProject = $true
            isEnableDue     = $true
            isEnableTag     = $true
        }
        evaluation   = @{
            isHideEvaluationSheet = $false
        }
        idle         = @{
            isEnableIdleTimeTracking      = $false
            minIdleTime                   = 60000
            isOnlyOpenIdleWhenCurrentTask = $false
        }
        takeABreak   = @{
            isTakeABreakEnabled            = $false
            isLockScreen                   = $false
            isTimedFullScreenBlocker       = $false
            timedFullScreenBlockerDuration = 5000
            isFocusWindow                  = $false
            takeABreakMessage              = "Take a break!"
            takeABreakMinWorkingTime       = 1800000
            takeABreakSnoozeTime           = 300000
            motivationalImgs               = @()
        }
        pomodoro     = @{
            cyclesBeforeLongerBreak = 4
        }
        keyboard     = @{}
        localBackup  = @{
            isEnabled = $false
        }
        sound        = @{
            isIncreaseDoneSoundPitch = $false
            doneSound                = $null
            breakReminderSound       = $null
            volume                   = 50
        }
        timeTracking = @{
            isAutoStartNextTask              = $false
            isNotifyWhenTimeEstimateExceeded = $false
            isTrackingReminderEnabled        = $false
            isTrackingReminderShowOnMobile   = $false
            trackingReminderMinTime          = 0
        }
        reminder     = @{
            isCountdownBannerEnabled = $false
            countdownDuration        = 60000
        }
        schedule     = @{
            isWorkStartEndEnabled = $false
            workStart             = "09:00"
            workEnd               = "17:00"
            isLunchBreakEnabled   = $false
            lunchBreakStart       = "12:00"
            lunchBreakEnd         = "13:00"
        }
        dominaMode   = @{
            isEnabled = $false
            text      = ""
            interval  = 10000
            volume    = 50
        }
        focusMode    = @{
            isSkipPreparation = $false
        }
        sync         = @{
            isEnabled    = $false
            syncProvider = $null
            syncInterval = 300000
        }
    }
}


# Process each list (convert to projects)
foreach ($list in $msTodo.lists) {
    $projectId = New-SPId

    # Create project entity
    $project = @{
        id               = $projectId
        title            = $list.displayName
        taskIds          = @()
        icon             = "list"  # Default icon # TODO: Map from MS Todo if possible
        isHiddenFromMenu = $false
        isArchived       = $false
        isEnableBacklog  = $false
        backlogTaskIds   = @()
        noteIds          = @()
        advancedCfg      = @{
            worklogExportSettings = @{
                cols            = @("DATE", "START", "END", "TIME_CLOCK", "TITLES_INCLUDING_SUB")
                separateTasksBy = " | "
                groupBy         = "DATE"
            }
        }
        theme            = @{
            isAutoContrast          = $true
            isDisableBackgroundTint = $false
            primary                 = "#42a5f5"
            huePrimary              = "500"
            accent                  = "#ff4081"
            hueAccent               = "500"
            warn                    = "#e11826"
            hueWarn                 = "500"
        }
    }

    # Add project to SP data
    $sp.project.ids += $projectId
    $sp.project.entities[$projectId] = $project

    # Add to imported projects
    $importedProjectIds += $projectId

    # Process tasks for this list
    if ($msTodo.tasks.PSObject.Properties[$list.id]) {
        foreach ($msTask in $msTodo.tasks.($list.id)) {
            $taskId = New-SPId

            # Parse timestamps
            # TODO: Handle potential nulls and parsing errors, extract parser to function
            $created = [DateTimeOffset]::Parse($msTask.createdDateTime).ToUnixTimeMilliseconds()
            $modified = [DateTimeOffset]::Parse($msTask.lastModifiedDateTime).ToUnixTimeMilliseconds()
            $doneOn = $null
            if ($msTask.status -eq "completed" -and $msTask.completedDateTime -and $msTask.completedDateTime.dateTime) {
                $doneOn = [DateTimeOffset]::Parse($msTask.completedDateTime.dateTime).ToUnixTimeMilliseconds()
            }

            # Collect tagIds
            $tagIds = @()
            if ($msTask.importance -eq "high") {
                $importantTagId = ($tagEntities.GetEnumerator() | Where-Object { $_.Value.title -eq "Important" }).Key
                if ($importantTagId) {
                    $tagIds += $importantTagId
                }
            }
            if ($msTask.categories) {
                foreach ($cat in $msTask.categories) {
                    $catTagId = ($tagEntities.GetEnumerator() | Where-Object { $_.Value.title -eq $cat }).Key
                    if ($catTagId) {
                        $tagIds += $catTagId
                    }
                }
            }

            # Create task entity
            $task = @{
                id             = $taskId
                title          = $msTask.title
                projectId      = $projectId
                isDone         = ($msTask.status -eq "completed")
                created        = $created
                modified       = $modified
                # Additional required fields
                subTaskIds     = @()
                tagIds         = $tagIds
                timeSpentOnDay = @{}
                timeEstimate   = 0
                timeSpent      = 0
                hasPlannedTime = $false
                attachments    = @()
            }

            # Add optional fields only if not null
            Add-IfNotNull $task "doneOn" $doneOn
            Add-IfNotNull $task "parentId" $null
            Add-IfNotNull $task "reminderId" $null
            Add-IfNotNull $task "repeatCfgId" $null
            Add-IfNotNull $task "dueWithTime" $null
            Add-IfNotNull $task "dueDay" $null
            Add-IfNotNull $task "remindAt" $null # TODO: Issue: remindAt is set to $null (line 354) but there's no code to parse reminderDateTime from MS Todo and map it to SP's reminder options
            Add-IfNotNull $task "_hideSubTasksMode" $null
            Add-IfNotNull $task "issueId" $null
            Add-IfNotNull $task "issueProviderId" $null
            Add-IfNotNull $task "issueType" $null
            Add-IfNotNull $task "issueWasUpdated" $null
            Add-IfNotNull $task "issueLastUpdated" $null
            Add-IfNotNull $task "issueAttachmentNr" $null
            Add-IfNotNull $task "issueTimeTracked" $null
            Add-IfNotNull $task "issuePoints" $null

            # Add notes if body content exists
            if ($msTask.body -and $msTask.body.content -and $msTask.body.content.Trim()) {
                $task["notes"] = $msTask.body.content.Trim()
            }

            # Handle recurrence
            if ($msTask.recurrence) {
                $repeatCfgId = New-SPId
                $startDateStr = Get-DbDateStr ([DateTimeOffset]::FromUnixTimeMilliseconds($task.created))
                $repeatCfg = @{
                    id                        = $repeatCfgId
                    projectId                 = $projectId
                    lastTaskCreation          = $task.created
                    lastTaskCreationDay       = $startDateStr
                    title                     = "Recurring Task"
                    tagIds                    = @()
                    order                     = 0
                    defaultEstimate           = 0
                    startTime                 = $null
                    remindAt                  = $null
                    isPaused                  = $false
                    quickSetting              = "DAILY"
                    repeatCycle               = "DAILY"
                    startDate                 = $null  # Optional, for monthly/yearly
                    repeatEvery               = 1
                    monday                    = $false
                    tuesday                   = $false
                    wednesday                 = $false
                    thursday                  = $false
                    friday                    = $false
                    saturday                  = $false
                    sunday                    = $false
                    notes                     = $null
                    shouldInheritSubtasks     = $false
                    repeatFromCompletionDate  = $false
                    disableAutoUpdateSubtasks = $false
                    subTaskTemplates          = @()
                    deletedInstanceDates      = @()
                }

                # Map recurrence pattern
                $pattern = $msTask.recurrence.pattern
                if ($pattern.type -eq "daily") {
                    $repeatCfg.repeatCycle = "DAILY"
                    $repeatCfg.repeatEvery = $pattern.interval
                    $repeatCfg.quickSetting = "DAILY"
                }
                elseif ($pattern.type -eq "weekly") {
                    $repeatCfg.repeatCycle = "WEEKLY"
                    $repeatCfg.repeatEvery = $pattern.interval
                    $repeatCfg.quickSetting = "WEEKLY_CURRENT_WEEKDAY" # TODO: Hardcoded without verification
                    # Set days if specified
                    if ($pattern.daysOfWeek) {
                        foreach ($day in $pattern.daysOfWeek) {
                            switch ($day) {
                                "monday" { $repeatCfg.monday = $true }
                                "tuesday" { $repeatCfg.tuesday = $true }
                                "wednesday" { $repeatCfg.wednesday = $true }
                                "thursday" { $repeatCfg.thursday = $true }
                                "friday" { $repeatCfg.friday = $true }
                                "saturday" { $repeatCfg.saturday = $true }
                                "sunday" { $repeatCfg.sunday = $true }
                            }
                        }
                    }
                    else {
                        # TODO: If no days specified, assume all days? But MS Todo weekly without days might be every week
                        # For simplicity, set to current day or something. Let's set monday for now.
                        $repeatCfg.monday = $true
                    }
                }
                elseif ($pattern.type -eq "monthly") {
                    $repeatCfg.repeatCycle = "MONTHLY"
                    $repeatCfg.repeatEvery = $pattern.interval
                    $repeatCfg.quickSetting = "MONTHLY_CURRENT_DATE"
                    $repeatCfg.startDate = $startDateStr
                }
                elseif ($pattern.type -eq "yearly") {
                    $repeatCfg.repeatCycle = "YEARLY"
                    $repeatCfg.repeatEvery = $pattern.interval
                    $repeatCfg.quickSetting = "YEARLY_CURRENT_DATE"
                    $repeatCfg.startDate = $startDateStr
                }

                # TODO: Map range - for endDate, we don't have endDate field in model, so perhaps set isPaused or something
                # The model doesn't have endDate, so maybe ignore or set deletedInstanceDates if needed
                # For now, leave as is

                $sp.taskRepeatCfg.entities[$repeatCfgId] = $repeatCfg
                $sp.taskRepeatCfg.ids += $repeatCfgId
                $task.repeatCfgId = $repeatCfgId
            }

            # Add task to SP data
            $sp.task.entities[$taskId] = $task

            # Add task ID to project's taskIds
            $project.taskIds += $taskId

            # Handle subtasks (checklistItems)
            if ($msTask.checklistItems) {
                foreach ($item in $msTask.checklistItems) {
                    $subTaskId = New-SPId
                    $subTask = @{
                        id             = $subTaskId
                        title          = $item.displayName
                        projectId      = $projectId
                        isDone         = $item.isChecked
                        created        = [DateTimeOffset]::Parse($item.createdDateTime).ToUnixTimeMilliseconds()
                        modified       = [DateTimeOffset]::Parse($item.createdDateTime).ToUnixTimeMilliseconds() # TODO: For subtasks, modified is set to the same value as created (line 470), assuming subtasks don't have separate modification times.
                        parentId       = $taskId
                        subTaskIds     = @()
                        tagIds         = @()
                        timeSpentOnDay = @{}
                        timeEstimate   = 0
                        timeSpent      = 0
                        hasPlannedTime = $false
                        attachments    = @()
                    }
                    $sp.task.entities[$subTaskId] = $subTask
                    $sp.task.ids += $subTaskId
                    $task.subTaskIds += $subTaskId
                }
            }
        }
    }
}

# Set task ids
$sp.task.ids = $sp.task.entities.Keys

# Build menuTree
if ($importedProjectIds.Count -gt 0) {
    $projectFolderId = [guid]::NewGuid().ToString()
    $projectChildren = @()
    foreach ($projId in $importedProjectIds) {
        $projectChildren += @{ k = "p"; id = $projId }
    }
    $sp.menuTree.projectTree = @(
        @{
            k          = "f"
            id         = $projectFolderId
            name       = "MS-IMPORT-PROJECTS"
            isExpanded = $true
            children   = $projectChildren
        }
    )
}

if ($importedTagIds.Count -gt 0) {
    $tagFolderId = [guid]::NewGuid().ToString()
    $tagChildren = @()
    foreach ($tid in $importedTagIds) {
        $tagChildren += @{ k = "t"; id = $tid }
    }
    $sp.menuTree.tagTree = @(
        @{
            k          = "f"
            id         = $tagFolderId
            name       = "MS-IMPORT-TAGS"
            isExpanded = $true
            children   = $tagChildren
        }
    )
}

# Convert to JSON and save
$spJson = $sp | ConvertTo-Json -Depth 10
$spJson | Out-File -FilePath $OutputFile -Encoding UTF8

Write-Host "Conversion complete. Output saved to $OutputFile"