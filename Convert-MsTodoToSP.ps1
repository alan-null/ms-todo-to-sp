# PowerShell script to convert Microsoft To Do JSON export to Super Productivity BACKUP format
#
# Supported input formats:
#   1. lists.json  – flat array of list objects, each with an embedded "tasks" array
#   2. ms-todo.json – object with { "lists": [...], "tasks": { "<listId>": [...] } }
#
# Output: CompleteBackup JSON ready to import into Super Productivity
#   (Settings → Sync & Backup → Import from File)
#
# Usage:
#   .\Convert-MsTodoToSP.ps1 -InputFile "temp\lists.json" -OutputFile "temp\sp_import.json"

param(
    [string]$InputFile = "temp\ms-todo.json",
    [string]$OutputFile = "temp\sp_converted.json"
)

Set-StrictMode -Off

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

# Generate a 21-character base64url ID (similar to SP's nanoid IDs)
function New-SPId {
    $bytes = [guid]::NewGuid().ToByteArray()
    $b64 = [Convert]::ToBase64String($bytes).Replace('+', '-').Replace('/', '_').TrimEnd('=')
    return $b64.Substring(0, [Math]::Min(21, $b64.Length))
}

# Return ISO date string YYYY-MM-DD from a DateTimeOffset
function Get-DbDateStr {
    param([DateTimeOffset]$date)
    return $date.ToString("yyyy-MM-dd")
}

# Return ISO date string YYYY-MM-DD for a specific day-of-month in a given year/month
function Get-DbDateStrFromDayOfMonth {
    param([DateTimeOffset]$baseDate, [int]$dayOfMonth)
    $maxDay = [DateTime]::DaysInMonth($baseDate.Year, $baseDate.Month)
    $d = [Math]::Min($dayOfMonth, $maxDay)
    return "$($baseDate.Year)-$($baseDate.Month.ToString('00'))-$($d.ToString('00'))"
}

# Try to parse a datetime string; return $null on failure / empty input
function Parse-Dto {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try { return [DateTimeOffset]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return $null }
}

# ---------------------------------------------------------------------------
# Static defaults matching SP's DEFAULT_PROJECT / DEFAULT_TAG constants
# ---------------------------------------------------------------------------

$NowMs = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()

$DefaultWorklogExport = [ordered]@{
    cols             = @("DATE", "START", "END", "TIME_CLOCK", "TITLES_INCLUDING_SUB")
    roundWorkTimeTo  = $null
    roundStartTimeTo = $null
    roundEndTimeTo   = $null
    separateTasksBy  = " | "
    groupBy          = "DATE"
}
$DefaultAdvancedCfg = [ordered]@{ worklogExportSettings = $DefaultWorklogExport }

$DefaultTagTheme = @{
    isAutoContrast           = $true
    isDisableBackgroundTint  = $false
    primary                  = "#a05db1"    # DEFAULT_TAG_COLOR
    huePrimary               = "500"
    accent                   = "#ff4081"
    hueAccent                = "500"
    warn                     = "#e11826"
    hueWarn                  = "500"
    backgroundImageDark      = $null
    backgroundImageLight     = $null
    backgroundOverlayOpacity = 20
}

$DefaultProjectTheme = @{
    isAutoContrast           = $true
    isDisableBackgroundTint  = $false
    primary                  = "#29a1aa"    # DEFAULT_PROJECT_COLOR
    huePrimary               = "500"
    accent                   = "#ff4081"
    hueAccent                = "500"
    warn                     = "#e11826"
    hueWarn                  = "500"
    backgroundImageDark      = $null
    backgroundImageLight     = $null
    backgroundOverlayOpacity = 20
}

$TodayTagTheme = @{
    isAutoContrast           = $true
    isDisableBackgroundTint  = $true
    primary                  = "#6495ED"    # DEFAULT_TODAY_TAG_COLOR
    huePrimary               = "400"
    accent                   = "#ff4081"
    hueAccent                = "500"
    warn                     = "#e11826"
    hueWarn                  = "500"
    backgroundImageDark      = ""
    backgroundImageLight     = $null
    backgroundOverlayOpacity = 20
}

# ---------------------------------------------------------------------------
# Parse input file (supports both formats)
# ---------------------------------------------------------------------------

$rawData = Get-Content -Path $InputFile -Encoding UTF8 | ConvertFrom-Json

$lists = @()
if ($rawData -is [array]) {
    # Format 1: flat array of lists with embedded tasks
    $lists = $rawData
}
elseif ($null -ne $rawData.lists) {
    # Format 2: { lists: [...], tasks: { listId: [...] } }
    foreach ($l in $rawData.lists) {
        $copy = $l | Select-Object *
        if ($null -ne $rawData.tasks -and
            $rawData.tasks.PSObject.Properties.Name -contains $l.id) {
            $copy | Add-Member -NotePropertyName 'tasks' `
                -NotePropertyValue $rawData.tasks.($l.id) `
                -Force
        }
        $lists += $copy
    }
}
else {
    Write-Error "Unrecognised input format. Expected array of lists or { lists, tasks } object."
    exit 1
}

# ---------------------------------------------------------------------------
# Pass 1 – collect all unique tag names from categories + importance
# ---------------------------------------------------------------------------

$tagNameSet = [System.Collections.Generic.SortedSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($list in $lists) {
    if (-not $list.tasks) { continue }
    foreach ($t in $list.tasks) {
        # Add "Important" tag only if at least one task actually needs it
        if ($t.importance -eq "high") {
            $tagNameSet.Add("Important") | Out-Null
        }
        if ($t.categories) {
            foreach ($c in $t.categories) {
                if (-not [string]::IsNullOrWhiteSpace($c)) {
                    $tagNameSet.Add($c.Trim()) | Out-Null
                }
            }
        }
        # Extract inline #tags from the task title (e.g. "Buy milk #shopping #urgent")
        if ($t.title -and $t.title -match '#') {
            foreach ($m in [regex]::Matches($t.title, '#(\w+)')) {
                $tagNameSet.Add($m.Groups[1].Value) | Out-Null
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Build tag entities (imported tags + mandatory TODAY_TAG)
# ---------------------------------------------------------------------------

$tagNameToId = @{}         # tagName -> tagId
$tagEntities = [ordered]@{}

# Imported tags first (they get added to the folder in menuTree)
foreach ($name in $tagNameSet) {
    $id = New-SPId
    $tagNameToId[$name] = $id
    $tagEntities[$id] = [ordered]@{
        id          = $id
        title       = $name
        color       = $null
        created     = $NowMs
        modified    = $NowMs
        icon        = $null
        taskIds     = [System.Collections.Generic.List[string]]::new()
        advancedCfg = $DefaultAdvancedCfg
        theme       = $DefaultTagTheme.Clone()
    }
}

# TODAY_TAG – a mandatory virtual tag; id must equal 'TODAY'
$TODAY_ID = 'TODAY'
$tagEntities[$TODAY_ID] = [ordered]@{
    id          = $TODAY_ID
    title       = 'Today'
    color       = $null
    created     = $NowMs
    modified    = $NowMs
    icon        = 'wb_sunny'
    taskIds     = [System.Collections.Generic.List[string]]::new()
    advancedCfg = $DefaultAdvancedCfg
    theme       = $TodayTagTheme.Clone()
}

# ---------------------------------------------------------------------------
# Processing state
# ---------------------------------------------------------------------------

$taskEntities = [ordered]@{}
$projectEntities = [ordered]@{}
$projectIds = [System.Collections.Generic.List[string]]::new()
$repeatCfgEntities = [ordered]@{}
$repeatCfgIds = [System.Collections.Generic.List[string]]::new()
$remindersList = [System.Collections.Generic.List[object]]::new()
$importedProjectIds = [System.Collections.Generic.List[string]]::new()
$repeatCfgOrder = 0

# ---------------------------------------------------------------------------
# Internal helper: map MS Graph recurrence pattern → SP repeat fields
# Returns a hashtable with SP repeat fields, or $null if pattern is unrecognised.
#
# MS Graph pattern types:
#   daily            – every N days
#   weekly           – every N weeks on daysOfWeek
#   absoluteMonthly  – every N months on dayOfMonth
#   relativeMonthly  – every N months on Nth weekday (e.g. "second Tuesday")
#                      → mapped to MONTHLY (best approximation, weekday constraint lost)
#   absoluteYearly   – every N years on month/dayOfMonth
#   relativeYearly   – every N years on Nth weekday of month
#                      → mapped to YEARLY (best approximation)
#   hourly           – every N hours → mapped to DAILY (closest SP option)
# ---------------------------------------------------------------------------
function Convert-RecurrencePattern {
    param(
        [object]$pattern,
        [DateTimeOffset]$createdDate,
        [string]$dueDateStr   # ISO datetime from dueDateTime (may be $null/empty)
    )

    if (-not $pattern) { return $null }

    $patType = if ($pattern.PSObject.Properties.Name -contains 'type') { $pattern.type } else { '' }
    $interval = if ($pattern.PSObject.Properties.Name -contains 'interval') { [int]($pattern.interval) } else { 1 }
    if ($interval -lt 1) { $interval = 1 }

    # Weekday flags stored in a hashtable to avoid nested-scope issues
    $days = @{ mon = $false; tue = $false; wed = $false; thu = $false; fri = $false; sat = $false; sun = $false }

    # Script block: set flags for each MS Graph day name (Title-case or lowercase)
    $setDays = {
        param([object[]]$dayNames)
        foreach ($d in $dayNames) {
            switch ($d.ToString().ToLower()) {
                "monday" { $days.mon = $true }
                "tuesday" { $days.tue = $true }
                "wednesday" { $days.wed = $true }
                "thursday" { $days.thu = $true }
                "friday" { $days.fri = $true }
                "saturday" { $days.sat = $true }
                "sunday" { $days.sun = $true }
            }
        }
    }

    # Script block: set the flag for whatever weekday $createdDate falls on
    $setCreatedDay = {
        switch ($createdDate.DayOfWeek) {
            ([DayOfWeek]::Monday) { $days.mon = $true }
            ([DayOfWeek]::Tuesday) { $days.tue = $true }
            ([DayOfWeek]::Wednesday) { $days.wed = $true }
            ([DayOfWeek]::Thursday) { $days.thu = $true }
            ([DayOfWeek]::Friday) { $days.fri = $true }
            ([DayOfWeek]::Saturday) { $days.sat = $true }
            ([DayOfWeek]::Sunday) { $days.sun = $true }
        }
    }

    $repeatCycle = $null
    $startDate = $null

    switch ($patType.ToLower()) {

        "daily" {
            $repeatCycle = "DAILY"
        }

        "hourly" {
            # SP has no hourly option; DAILY is the closest approximation
            $repeatCycle = "DAILY"
            $interval = 1
            Write-Warning "  [recurrence] 'hourly' pattern mapped to DAILY (SP has no hourly repeat)."
        }

        "weekly" {
            $repeatCycle = "WEEKLY"
            $hasDays = ($pattern.PSObject.Properties.Name -contains 'daysOfWeek') -and
            $pattern.daysOfWeek -and
            $pattern.daysOfWeek.Count -gt 0
            if ($hasDays) { & $setDays $pattern.daysOfWeek }
            else { & $setCreatedDay }
        }

        { $_ -eq "absolutemonthly" -or $_ -eq "relativemonthly" } {
            $repeatCycle = "MONTHLY"

            # Prefer dueDate's day; fall back to dayOfMonth field; then createdDate
            if (-not [string]::IsNullOrWhiteSpace($dueDateStr)) {
                $dto = Parse-Dto $dueDateStr
                if ($dto) { $startDate = Get-DbDateStr $dto }
            }
            if (-not $startDate -and
                $pattern.PSObject.Properties.Name -contains 'dayOfMonth' -and
                $pattern.dayOfMonth -gt 0) {
                $startDate = Get-DbDateStrFromDayOfMonth $createdDate ([int]$pattern.dayOfMonth)
            }
            if (-not $startDate) { $startDate = Get-DbDateStr $createdDate }

            if ($patType.ToLower() -eq "relativemonthly") {
                Write-Warning "  [recurrence] 'relativeMonthly' (Nth weekday of month) mapped to MONTHLY – weekday constraint lost."
            }
        }

        { $_ -eq "absoluteyearly" -or $_ -eq "relativeyearly" } {
            $repeatCycle = "YEARLY"

            if (-not [string]::IsNullOrWhiteSpace($dueDateStr)) {
                $dto = Parse-Dto $dueDateStr
                if ($dto) { $startDate = Get-DbDateStr $dto }
            }
            if (-not $startDate) { $startDate = Get-DbDateStr $createdDate }

            if ($patType.ToLower() -eq "relativeyearly") {
                Write-Warning "  [recurrence] 'relativeYearly' (Nth weekday of month/year) mapped to YEARLY – weekday constraint lost."
            }
        }

        default {
            Write-Warning "  [recurrence] Unknown pattern type '$patType' – skipping repeat config."
            return $null
        }
    }

    # Determine quickSetting
    $quickSetting = "CUSTOM"
    $activeDays = @($days.mon, $days.tue, $days.wed, $days.thu, $days.fri, $days.sat, $days.sun) | Where-Object { $_ }
    $weekdayCount = @($days.mon, $days.tue, $days.wed, $days.thu, $days.fri) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count
    if ($repeatCycle -eq "DAILY" -and $interval -eq 1) {
        $quickSetting = "DAILY"
    }
    elseif ($repeatCycle -eq "WEEKLY" -and $interval -eq 1) {
        if ($activeDays.Count -eq 1) {
            $quickSetting = "WEEKLY_CURRENT_WEEKDAY"
        }
        elseif ($weekdayCount -eq 5 -and -not $days.sat -and -not $days.sun) {
            $quickSetting = "MONDAY_TO_FRIDAY"
        }
    }
    elseif ($repeatCycle -eq "MONTHLY" -and $interval -eq 1) {
        $quickSetting = "MONTHLY_CURRENT_DATE"
    }
    elseif ($repeatCycle -eq "YEARLY" -and $interval -eq 1) {
        $quickSetting = "YEARLY_CURRENT_DATE"
    }

    return [ordered]@{
        repeatCycle  = $repeatCycle
        repeatEvery  = $interval
        quickSetting = $quickSetting
        monday       = $days.mon
        tuesday      = $days.tue
        wednesday    = $days.wed
        thursday     = $days.thu
        friday       = $days.fri
        saturday     = $days.sat
        sunday       = $days.sun
        startDate    = $startDate   # $null for DAILY/WEEKLY; string for MONTHLY/YEARLY
    }
}

# ---------------------------------------------------------------------------
# Pass 2 – convert lists → projects and tasks → SP tasks
# ---------------------------------------------------------------------------

foreach ($list in $lists) {
    $projectId = New-SPId

    $project = [ordered]@{
        id               = $projectId
        title            = $list.displayName
        taskIds          = [System.Collections.Generic.List[string]]::new()
        icon             = "list_alt"
        isHiddenFromMenu = $false
        isArchived       = $false
        isEnableBacklog  = $false
        backlogTaskIds   = @()
        noteIds          = @()
        advancedCfg      = $DefaultAdvancedCfg
        theme            = $DefaultProjectTheme.Clone()
    }

    $projectEntities[$projectId] = $project
    $projectIds.Add($projectId)
    $importedProjectIds.Add($projectId)

    if (-not $list.tasks) { continue }

    foreach ($msTask in $list.tasks) {
        if ([string]::IsNullOrWhiteSpace($msTask.title)) { continue }

        $taskId = New-SPId

        # ---- extract inline #tags from title for linking (title is kept as-is) ----
        $inlineTagMatches = [regex]::Matches($msTask.title, '#(\w+)')

        # --- timestamps ---
        $created = $NowMs
        $modified = $NowMs
        $dto = Parse-Dto $msTask.createdDateTime
        if ($dto) { $created = $dto.ToUnixTimeMilliseconds() }
        $dto = Parse-Dto $msTask.lastModifiedDateTime
        if ($dto) { $modified = $dto.ToUnixTimeMilliseconds() }

        $isDone = ($msTask.status -eq "completed")
        $doneOn = $null
        if ($isDone -and
            $msTask.PSObject.Properties.Name -contains 'completedDateTime' -and
            $msTask.completedDateTime -and
            $msTask.completedDateTime.PSObject.Properties.Name -contains 'dateTime') {
            $dto = Parse-Dto $msTask.completedDateTime.dateTime
            if ($dto) { $doneOn = $dto.ToUnixTimeMilliseconds() }
        }
        # Fallback: if task is done but no completedDateTime, use lastModified
        if ($isDone -and $null -eq $doneOn) { $doneOn = $modified }

        # --- tag IDs ---
        $taskTagIds = [System.Collections.Generic.List[string]]::new()

        if ($msTask.importance -eq "high") {
            $tid = $tagNameToId["Important"]
            if ($tid) {
                $taskTagIds.Add($tid)
                if (-not $tagEntities[$tid].taskIds.Contains($taskId)) {
                    $tagEntities[$tid].taskIds.Add($taskId)
                }
            }
        }

        if ($msTask.categories) {
            foreach ($cat in $msTask.categories) {
                if ([string]::IsNullOrWhiteSpace($cat)) { continue }
                $tid = $tagNameToId[$cat.Trim()]
                if ($tid -and -not $taskTagIds.Contains($tid)) {
                    $taskTagIds.Add($tid)
                    if (-not $tagEntities[$tid].taskIds.Contains($taskId)) {
                        $tagEntities[$tid].taskIds.Add($taskId)
                    }
                }
            }
        }

        # ---- inline #tags extracted from title --------------------------------
        foreach ($m in $inlineTagMatches) {
            $tid = $tagNameToId[$m.Groups[1].Value]
            if ($tid -and -not $taskTagIds.Contains($tid)) {
                $taskTagIds.Add($tid)
                if (-not $tagEntities[$tid].taskIds.Contains($taskId)) {
                    $tagEntities[$tid].taskIds.Add($taskId)
                }
            }
        }

        # ---- due date ---------------------------------------------------------
        # dueWithTime and dueDay are mutually exclusive; check dueWithTime first.
        $dueWithTime = $null
        $dueDay = $null
        $dueDateIsoStr = $null   # raw ISO string kept for recurrence startDate derivation

        if ($msTask.PSObject.Properties.Name -contains 'dueDateTime' -and
            $msTask.dueDateTime -and
            $msTask.dueDateTime.PSObject.Properties.Name -contains 'dateTime') {
            $dto = Parse-Dto $msTask.dueDateTime.dateTime
            if ($dto) {
                $dueDateIsoStr = $msTask.dueDateTime.dateTime
                # If time-of-day component is less than 1 minute → treat as date-only
                if ($dto.TimeOfDay.TotalSeconds -lt 60) {
                    $dueDay = Get-DbDateStr $dto
                }
                else {
                    $dueWithTime = $dto.ToUnixTimeMilliseconds()
                }
            }
        }
        $hasPlannedTime = ($null -ne $dueDay -or $null -ne $dueWithTime)

        # --- reminder (reminderDateTime) ---
        $reminderId = $null
        $remindAt = $null
        if ($msTask.isReminderOn -eq $true -and
            $msTask.PSObject.Properties.Name -contains 'reminderDateTime' -and
            $msTask.reminderDateTime -and
            $msTask.reminderDateTime.PSObject.Properties.Name -contains 'dateTime') {
            $dto = Parse-Dto $msTask.reminderDateTime.dateTime
            if ($dto) {
                $remindAt = $dto.ToUnixTimeMilliseconds()
                $reminderId = New-SPId
                $remindersList.Add([ordered]@{
                        id        = $reminderId
                        remindAt  = $remindAt
                        title     = $msTask.title
                        type      = "TASK"
                        relatedId = $taskId
                    })
            }
        }

        # --- build task entity ---
        $task = [ordered]@{
            id             = $taskId
            title          = $msTask.title
            projectId      = $projectId
            isDone         = $isDone
            created        = $created
            modified       = $modified
            subTaskIds     = [System.Collections.Generic.List[string]]::new()
            tagIds         = @($taskTagIds)
            timeSpentOnDay = [ordered]@{}
            timeEstimate   = 0
            timeSpent      = 0
            hasPlannedTime = $hasPlannedTime
            attachments    = @()
        }

        if ($null -ne $doneOn) { $task["doneOn"] = $doneOn }
        if ($null -ne $dueWithTime) { $task["dueWithTime"] = $dueWithTime }
        if ($null -ne $dueDay) { $task["dueDay"] = $dueDay }
        if ($null -ne $reminderId) {
            $task["reminderId"] = $reminderId
            $task["remindAt"] = $remindAt
        }

        # --- notes (body) ---
        if ($msTask.PSObject.Properties.Name -contains 'body' -and
            $msTask.body -and
            $msTask.body.PSObject.Properties.Name -contains 'content' -and
            -not [string]::IsNullOrWhiteSpace($msTask.body.content)) {
            $task["notes"] = $msTask.body.content.Trim()
        }

        # ---- recurrence -------------------------------------------------------
        # Completed tasks are not migrated with a repeat config – they are
        # historical instances and the repeat should start fresh in SP.
        if (-not $isDone -and
            $msTask.PSObject.Properties.Name -contains 'recurrence' -and
            $msTask.recurrence -and
            $msTask.recurrence.PSObject.Properties.Name -contains 'pattern' -and
            $msTask.recurrence.pattern) {

            $taskCreatedDate = [DateTimeOffset]::FromUnixTimeMilliseconds($created)
            $startDateStr = Get-DbDateStr $taskCreatedDate

            $rp = Convert-RecurrencePattern `
                -pattern     $msTask.recurrence.pattern `
                -createdDate $taskCreatedDate `
                -dueDateStr  $dueDateIsoStr

            if ($rp) {
                $repeatCfgId = New-SPId
                $repeatNotes = if ($task.Contains("notes")) { $task["notes"] } else { $null }

                # Snapshot tagIds array now (taskTagIds is still a List at this point)
                $repeatTagIds = @($taskTagIds)

                $repeatCfg = [ordered]@{
                    id                       = $repeatCfgId
                    projectId                = $projectId
                    lastTaskCreation         = $created
                    lastTaskCreationDay      = $startDateStr
                    title                    = $msTask.title
                    tagIds                   = $repeatTagIds
                    order                    = $repeatCfgOrder++
                    isPaused                 = $false
                    quickSetting             = $rp.quickSetting
                    repeatCycle              = $rp.repeatCycle
                    repeatEvery              = $rp.repeatEvery
                    monday                   = $rp.monday
                    tuesday                  = $rp.tuesday
                    wednesday                = $rp.wednesday
                    thursday                 = $rp.thursday
                    friday                   = $rp.friday
                    saturday                 = $rp.saturday
                    sunday                   = $rp.sunday
                    notes                    = $repeatNotes
                    shouldInheritSubtasks    = $false
                    repeatFromCompletionDate = $false
                    deletedInstanceDates     = @()
                }

                # startDate is only meaningful for MONTHLY / YEARLY
                if ($rp.repeatCycle -eq "MONTHLY" -or $rp.repeatCycle -eq "YEARLY") {
                    $repeatCfg["startDate"] = if ($rp.startDate) { $rp.startDate } else { $startDateStr }
                }

                $repeatCfgEntities[$repeatCfgId] = $repeatCfg
                $repeatCfgIds.Add($repeatCfgId)
                $task["repeatCfgId"] = $repeatCfgId
            }
        }

        # Register parent task
        $taskEntities[$taskId] = $task
        $project.taskIds.Add($taskId) | Out-Null

        # --- subtasks (checklistItems) ---
        if ($msTask.PSObject.Properties.Name -contains 'checklistItems' -and
            $msTask.checklistItems) {
            foreach ($item in $msTask.checklistItems) {
                if ([string]::IsNullOrWhiteSpace($item.displayName)) { continue }
                $subId = New-SPId
                $subCreated = $NowMs
                $dto = Parse-Dto $item.createdDateTime
                if ($dto) { $subCreated = $dto.ToUnixTimeMilliseconds() }

                $subTask = [ordered]@{
                    id             = $subId
                    title          = $item.displayName
                    projectId      = $projectId
                    parentId       = $taskId
                    isDone         = [bool]$item.isChecked
                    created        = $subCreated
                    modified       = $subCreated
                    subTaskIds     = @()
                    tagIds         = @()
                    timeSpentOnDay = [ordered]@{}
                    timeEstimate   = 0
                    timeSpent      = 0
                    hasPlannedTime = $false
                    attachments    = @()
                }
                $taskEntities[$subId] = $subTask
                $task.subTaskIds.Add($subId) | Out-Null
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Finalise: convert List<string> → plain arrays for JSON serialisation
# ---------------------------------------------------------------------------

foreach ($id in @($taskEntities.Keys)) {
    $t = $taskEntities[$id]
    if ($t.subTaskIds -is [System.Collections.Generic.List[string]]) {
        $t.subTaskIds = @($t.subTaskIds)
    }
    if ($t.tagIds -is [System.Collections.Generic.List[string]]) {
        $t.tagIds = @($t.tagIds)
    }
}

foreach ($id in @($projectIds)) {
    if ($projectEntities[$id].taskIds -is [System.Collections.Generic.List[string]]) {
        $projectEntities[$id].taskIds = @($projectEntities[$id].taskIds)
    }
}

foreach ($id in @($tagEntities.Keys)) {
    if ($tagEntities[$id].taskIds -is [System.Collections.Generic.List[string]]) {
        $tagEntities[$id].taskIds = @($tagEntities[$id].taskIds)
    }
}

# Ensure repeatCfg tagIds are plain arrays
foreach ($id in @($repeatCfgIds)) {
    $rc = $repeatCfgEntities[$id]
    if ($rc.tagIds -is [System.Collections.Generic.List[string]]) {
        $rc.tagIds = @($rc.tagIds)
    }
}

# Build ordered tag ids: imported tags alphabetically, then TODAY_TAG last
# (TODAY always exists in the state; its position in menuTree puts it first)
$tagIds = [System.Collections.Generic.List[string]]::new()
foreach ($id in @($tagEntities.Keys)) {
    if ($id -ne $TODAY_ID) { $tagIds.Add($id) }
}
$tagIds.Add($TODAY_ID)

# ---------------------------------------------------------------------------
# Build menuTree
# Using k="p" for project node, k="t" for tag node, k="f" for folder
# ---------------------------------------------------------------------------

$projectTree = @()
if ($importedProjectIds.Count -gt 0) {
    $children = @()
    foreach ($_pid in $importedProjectIds) { $children += @{ k = "p"; id = $_pid } }
    $projectTree = @(@{
            k          = "f"
            id         = [guid]::NewGuid().ToString()
            name       = "MS-Imported"
            isExpanded = $true
            children   = $children
        })
}

# Tag tree: TODAY_TAG at root, imported tags in a folder
$tagTree = @(@{ k = "t"; id = $TODAY_ID })
$importedTagNodes = @()
foreach ($id in @($tagIds)) {
    if ($id -ne $TODAY_ID) { $importedTagNodes += @{ k = "t"; id = $id } }
}
if ($importedTagNodes.Count -gt 0) {
    $tagTree += @{
        k          = "f"
        id         = [guid]::NewGuid().ToString()
        name       = "MS-Imported Tags"
        isExpanded = $true
        children   = $importedTagNodes
    }
}

# ---------------------------------------------------------------------------
# Assemble AppDataComplete
# ---------------------------------------------------------------------------

$appData = [ordered]@{
    task           = [ordered]@{
        ids                   = @($taskEntities.Keys)
        entities              = $taskEntities
        currentTaskId         = $null
        selectedTaskId        = $null
        taskDetailTargetPanel = $null
        lastCurrentTaskId     = $null
        isDataLoaded          = $true
    }
    project        = [ordered]@{
        ids      = @($projectIds)
        entities = $projectEntities
    }
    tag            = [ordered]@{
        ids      = @($tagIds)
        entities = $tagEntities
    }
    taskRepeatCfg  = [ordered]@{
        ids      = @($repeatCfgIds)
        entities = $repeatCfgEntities
    }
    simpleCounter  = [ordered]@{
        ids      = @()
        entities = [ordered]@{}
    }
    metric         = [ordered]@{
        ids      = @()
        entities = [ordered]@{}
    }
    reminders      = @($remindersList)
    planner        = [ordered]@{
        days                           = [ordered]@{}
        addPlannedTasksDialogLastShown = $null
    }
    boards         = [ordered]@{
        boardCfgs = @()
    }
    note           = [ordered]@{
        ids        = @()
        entities   = [ordered]@{}
        todayOrder = @()
    }
    issueProvider  = [ordered]@{
        ids      = @()
        entities = [ordered]@{}
    }
    menuTree       = [ordered]@{
        projectTree = $projectTree
        tagTree     = $tagTree
    }
    timeTracking   = [ordered]@{
        project = [ordered]@{}
        tag     = [ordered]@{}
    }
    archiveYoung   = [ordered]@{
        task                  = [ordered]@{ ids = @(); entities = [ordered]@{} }
        timeTracking          = [ordered]@{ project = [ordered]@{}; tag = [ordered]@{} }
        lastTimeTrackingFlush = 0
    }
    archiveOld     = [ordered]@{
        task                  = [ordered]@{ ids = @(); entities = [ordered]@{} }
        timeTracking          = [ordered]@{ project = [ordered]@{}; tag = [ordered]@{} }
        lastTimeTrackingFlush = 0
    }
    pluginMetadata = @()
    pluginUserData = @()
    # Minimal globalConfig – the app's data-repair will fill in any missing keys
    globalConfig   = [ordered]@{
        misc         = [ordered]@{
            isConfirmBeforeExit                 = $false
            isConfirmBeforeExitWithoutFinishDay = $false
            isMinimizeToTray                    = $false
            startOfNextDay                      = 0
            isDisableAnimations                 = $false
        }
        tasks        = [ordered]@{
            isAutoMarkParentAsDone             = $false
            isMarkdownFormattingInNotesEnabled = $true
            notesTemplate                      = ""
        }
        shortSyntax  = [ordered]@{
            isEnableProject = $true
            isEnableDue     = $true
            isEnableTag     = $true
        }
        evaluation   = [ordered]@{ isHideEvaluationSheet = $false }
        idle         = [ordered]@{
            isEnableIdleTimeTracking      = $false
            minIdleTime                   = 60000
            isOnlyOpenIdleWhenCurrentTask = $false
        }
        takeABreak   = [ordered]@{
            isTakeABreakEnabled      = $false
            takeABreakMessage        = "Take a break!"
            takeABreakMinWorkingTime = 1800000
            motivationalImgs         = @()
        }
        pomodoro     = [ordered]@{ cyclesBeforeLongerBreak = 4 }
        sound        = [ordered]@{
            isIncreaseDoneSoundPitch = $false
            doneSound                = $null
            breakReminderSound       = $null
            volume                   = 50
        }
        timeTracking = [ordered]@{
            isAutoStartNextTask              = $false
            isNotifyWhenTimeEstimateExceeded = $false
            isTrackingReminderEnabled        = $false
            trackingReminderMinTime          = 0
        }
        schedule     = [ordered]@{
            isWorkStartEndEnabled = $false
            workStart             = "09:00"
            workEnd               = "17:00"
        }
        sync         = [ordered]@{
            isEnabled    = $false
            syncProvider = $null
            syncInterval = 300000
        }
    }
}

# ---------------------------------------------------------------------------
# Wrap in CompleteBackup format (recognized by importCompleteBackup())
# ---------------------------------------------------------------------------

$output = [ordered]@{
    timestamp         = $NowMs
    lastUpdate        = $NowMs
    crossModelVersion = 4.5
    data              = $appData
}

# ---------------------------------------------------------------------------
# Serialize and write output
# ---------------------------------------------------------------------------

$json = $output | ConvertTo-Json -Depth 20

$resolvedOutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)

# Ensure parent directory exists
$outDir = Split-Path $resolvedOutputFile -Parent
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

[System.IO.File]::WriteAllText(
    $resolvedOutputFile,
    $json,
    [System.Text.Encoding]::UTF8
)

$taskCount = @($taskEntities.Keys).Count
$projectCount = $projectIds.Count
$tagCount = $tagIds.Count - 1   # exclude TODAY_TAG
$repeatCount = $repeatCfgIds.Count
$reminderCount = $remindersList.Count

Write-Host "Conversion complete."
Write-Host "  Projects : $projectCount"
Write-Host "  Tasks    : $taskCount (incl. subtasks)"
Write-Host "  Tags     : $tagCount (+ TODAY tag)"
Write-Host "  Repeats  : $repeatCount"
Write-Host "  Reminders: $reminderCount"
Write-Host "  Output   : $resolvedOutputFile"
