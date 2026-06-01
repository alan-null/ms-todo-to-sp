# Merge-SPBackups.ps1
# ---------------------------------------------------------------------------
# Merges two Super Productivity backup files with deduplication and conflict
# resolution.  The Primary file is the "base" - its data is preserved, and
# unique data from the Secondary file is merged in.
#
# Duplicate detection:
#   - Projects    - by ID, then by title (case-insensitive)
#   - Tags        - by ID (incl. well-known "TODAY"), then by title
#   - Tasks       - by ID, then by title within the same (mapped) project
#                    only when instance anchors (dates/timestamps) match
#   - RepeatCfgs  - by ID, then by title within the same (mapped) project
#   - Reminders   - by mapped relatedId
#
# Conflict handling for tasks (NON-DESTRUCTIVE):
#   - Nothing in the Primary file is ever overwritten, merged into, or deleted.
#   - When a Secondary task conflicts with a Primary one (same ID, or same
#     title within the same mapped project AND a matching instance anchor), the
#     Secondary task is kept as its own separate entry and its title is prefixed
#     with "[CONFLICT] " so it can be reconciled by hand later.
#   - If the conflict is an ID collision, the Secondary copy is given a fresh
#     unique ID so both tasks can coexist; parentId / subTaskIds references are
#     rewritten to match.
#   - Non-conflicting Secondary tasks are added normally (unprefixed).
#   - Projects, tags and repeat configs are still de-duplicated by ID/title;
#     this never removes a task, it only routes Secondary tasks into the
#     matching Primary container.
#
# Usage:
#   .\Merge-SPBackups.ps1 -PrimaryFile "a.json" -SecondaryFile "b.json" -OutputFile "merged.json"
# ---------------------------------------------------------------------------

param(
    [string]$PrimaryFile    = "input\sp-backup.json",
    [string]$SecondaryFile  = "out\ms-converted-alpl.json",
    [string]$OutputFile     = "out\sp-and-ms.json"
)

Set-StrictMode -Off

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Parse a JSON string into a nested ordered-hashtable tree, preserving empty
# arrays ([]) as @() and single-element arrays (["x"]) as @("x").
# ConvertFrom-Json (PS5.1) collapses both of those to $null / a bare scalar.
function ConvertFrom-JsonPreserveArrays {
    param([string]$rawJson)
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        # PS7 / .NET 5+ — System.Text.Json is built in
        function Convert-JsonElement {
            param($el)
            switch ($el.ValueKind.ToString()) {
                'Object' {
                    $ht = [ordered]@{}
                    foreach ($p in $el.EnumerateObject()) {
                        $ht[$p.Name] = Convert-JsonElement $p.Value
                    }
                    return $ht
                }
                'Array' {
                    $arr = @($el.EnumerateArray() | ForEach-Object { Convert-JsonElement $_ })
                    Write-Output -NoEnumerate $arr
                    return
                }
                'String' { return $el.GetString() }
                'Number' {
                    $i = [long]0
                    if ($el.TryGetInt64([ref]$i)) { return $i }
                    return $el.GetDouble()
                }
                'True'  { return $true }
                'False' { return $false }
                default { return $null }
            }
        }
        $doc = [System.Text.Json.JsonDocument]::Parse($rawJson)
        try { return Convert-JsonElement $doc.RootElement } finally { $doc.Dispose() }
    }
    else {
        # PS5.1 / .NET Framework — use JavaScriptSerializer
        Add-Type -AssemblyName System.Web.Extensions
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = [int]::MaxValue
        function Convert-JsObject {
            param($obj)
            if ($null -eq $obj) { return $null }
            if ($obj -is [System.Collections.Generic.Dictionary[string,object]]) {
                $ht = [ordered]@{}
                foreach ($k in $obj.Keys) { $ht[$k] = Convert-JsObject $obj[$k] }
                return $ht
            }
            if ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
                $arr = @($obj | ForEach-Object { Convert-JsObject $_ })
                Write-Output -NoEnumerate $arr
                return
            }
            return $obj
        }
        return Convert-JsObject ($ser.DeserializeObject($rawJson))
    }
}

# Ensure a parsed backup is in the canonical "wrapped" shape:
#   { timestamp; lastUpdate; crossModelVersion; data = { task; project; tag; ... } }
#
# Super Productivity backups come in two on-disk shapes:
#   * Full backup  - wrapped: a top-level "data" object holds the model, with
#                    timestamp / lastUpdate / crossModelVersion alongside it.
#   * Data export  - flat:    the model (task / project / tag / ...) sits at the
#                    very top level, with no "data" wrapper and no metadata.
# Every merge operation below addresses "$x.data.*", so a flat file must be
# wrapped first; otherwise "$x.data" is $null and all access cascades into
# NullArray / InvokeMethodOnNull / PropertyNotFound errors.
function ConvertTo-WrappedBackup {
    param($parsed)
    if ($null -eq $parsed) { return $parsed }

    $isDict = $parsed -is [System.Collections.IDictionary]

    # Already wrapped: has a "data" member that is itself an object.
    if ($isDict -and $parsed.Contains("data") -and ($parsed.data -is [System.Collections.IDictionary])) {
        return $parsed
    }

    # Flat model -> wrap it, carrying over any top-level metadata if present.
    $wrapped = [ordered]@{}
    $wrapped["timestamp"]         = if ($isDict -and $parsed.Contains("timestamp"))         { $parsed["timestamp"] }         else { 0 }
    $wrapped["lastUpdate"]        = if ($isDict -and $parsed.Contains("lastUpdate"))        { $parsed["lastUpdate"] }        else { 0 }
    $wrapped["crossModelVersion"] = if ($isDict -and $parsed.Contains("crossModelVersion")) { $parsed["crossModelVersion"] } else { 4.5 }
    $wrapped["data"]              = $parsed
    return $wrapped
}

# Normalise a title for comparison
function Normalize-Title {
    param([string]$t)
    if (-not $t) { return "" }
    return $t.Trim().ToLower()
}

# Safe array-add (returns new array with $val appended if not present)
function Safe-ArrayAdd {
    param([object]$arr, [string]$val)
    if ($null -eq $arr) { $arr = @() }
    if ($val -in $arr) { Write-Output -NoEnumerate $arr; return }
    $result = @($arr) + @($val)
    Write-Output -NoEnumerate $result
}

# Prefix a secondary task title to flag an unresolved conflict (idempotent).
function Add-ConflictPrefix {
    param([string]$title)
    if ($null -eq $title) { $title = "" }
    if ($title.StartsWith("[CONFLICT]")) { return $title }
    return "[CONFLICT] $title"
}

# Mint a task id that is not used by either backup or a previously minted id.
# Used only when a secondary task id collides with a primary one, so both copies
# can coexist (the primary task is never overwritten).
$script:usedTaskIds = @{}
$script:idAlphabet  = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
function New-UniqueTaskId {
    do {
        $chars = for ($i = 0; $i -lt 21; $i++) { $script:idAlphabet[(Get-Random -Maximum $script:idAlphabet.Length)] }
        $id = -join $chars
    } while ($script:usedTaskIds.Contains($id))
    $script:usedTaskIds[$id] = $true
    return $id
}

# Pick the best task match from a list of candidates by closest created timestamp
function Find-BestTaskMatch {
    param([object[]]$candidates, $sourceTask)
    if ($candidates.Count -eq 0) { return $null }
    if ($candidates.Count -eq 1) { return $candidates[0] }
    $srcCreated = if ($sourceTask.Contains("created")) { $sourceTask.created } else { 0 }
    $best = $null; $bestDiff = [long]::MaxValue
    foreach ($c in $candidates) {
        $cc = if ($c.Contains("created")) { $c.created } else { 0 }
        $diff = [Math]::Abs($cc - $srcCreated)
        if ($diff -lt $bestDiff) { $bestDiff = $diff; $best = $c }
    }
    return $best
}

# Determine if two tasks likely represent the same task instance (not only same title)
function Is-SameTaskInstance {
    param($priTask, $secTask, [string]$mappedSecRepeatCfgId = $null)

    $priRepeat = if ($priTask.Contains("repeatCfgId")) { $priTask.repeatCfgId } else { $null }
    $secRepeat = if ($mappedSecRepeatCfgId) {
        $mappedSecRepeatCfgId
    }
    elseif ($secTask.Contains("repeatCfgId")) {
        $secTask.repeatCfgId
    }
    else {
        $null
    }

    if (($priRepeat -and -not $secRepeat) -or (-not $priRepeat -and $secRepeat)) { return $false }
    if ($priRepeat -and $secRepeat -and $priRepeat -ne $secRepeat) { return $false }

    foreach ($f in @("created", "doneOn", "dueDay", "dueWithTime", "remindAt")) {
        $pHas = $priTask.Contains($f) -and $null -ne $priTask[$f] -and "$($priTask[$f])" -ne "" -and $priTask[$f] -ne 0
        $sHas = $secTask.Contains($f) -and $null -ne $secTask[$f] -and "$($secTask[$f])" -ne "" -and $secTask[$f] -ne 0
        if ($pHas -and $sHas -and $priTask[$f] -eq $secTask[$f]) {
            return $true
        }
    }

    return $false
}

# ---------------------------------------------------------------------------
# Merge helpers
# ---------------------------------------------------------------------------

# Merge individual task fields from $sec into $pri (in-place).
# Returns $true if any field was updated.
function Merge-TaskFields {
    param($priTask, $secTask)
    $updated = $false

    # --- status / dates: newer "modified" wins ---
    $secMod = if ($secTask.Contains("modified")) { $secTask.modified } else { 0 }
    $priMod = if ($priTask.Contains("modified")) { $priTask.modified } else { 0 }
    if ($secMod -gt $priMod) {
        $priTask.modified = $secMod
        $priTask.isDone = $secTask.isDone
        foreach ($f in @("doneOn", "dueDay", "dueWithTime")) {
            if ($secTask.Contains($f)) { $priTask[$f] = $secTask[$f]; $updated = $true }
        }
        $priTask.hasPlannedTime = $secTask.hasPlannedTime
        $updated = $true
    }

    # --- notes: append if different ---
    $secNotes = if ($secTask.Contains("notes")) { $secTask.notes } else { $null }
    $priNotes = if ($priTask.Contains("notes")) { $priTask.notes } else { $null }
    if ($secNotes -and -not [string]::IsNullOrWhiteSpace($secNotes)) {
        if ([string]::IsNullOrWhiteSpace($priNotes)) {
            $priTask["notes"] = $secNotes; $updated = $true
        }
        elseif ($priNotes.Trim() -ne $secNotes.Trim()) {
            $priTask["notes"] = "$priNotes`n---`n$secNotes"; $updated = $true
        }
    }

    # --- timeSpentOnDay: max per day ---
    if ($secTask.Contains("timeSpentOnDay") -and $secTask.timeSpentOnDay) {
        if (-not $priTask.Contains("timeSpentOnDay") -or $null -eq $priTask.timeSpentOnDay) {
            $priTask["timeSpentOnDay"] = [ordered]@{}
        }
        foreach ($day in $secTask.timeSpentOnDay.Keys) {
            $sv = $secTask.timeSpentOnDay[$day]
            if (-not $priTask.timeSpentOnDay.Contains($day)) {
                $priTask.timeSpentOnDay[$day] = $sv; $updated = $true
            }
            elseif ($sv -gt $priTask.timeSpentOnDay[$day]) {
                $priTask.timeSpentOnDay[$day] = $sv; $updated = $true
            }
        }
    }

    # --- time estimates / spent: take max ---
    foreach ($f in @("timeEstimate", "timeSpent")) {
        $sv = if ($secTask.Contains($f)) { $secTask[$f] } else { 0 }
        $pv = if ($priTask.Contains($f)) { $priTask[$f] } else { 0 }
        if ($sv -gt $pv) { $priTask[$f] = $sv; $updated = $true }
    }

    # --- remindAt: prefer whichever is set (or newer) ---
    if ($secTask.Contains("remindAt") -and $secTask.remindAt) {
        if (-not $priTask.Contains("remindAt") -or -not $priTask.remindAt) {
            $priTask["remindAt"] = $secTask.remindAt; $updated = $true
        }
    }

    return $updated
}

# ===========================================================================
Write-Host "============================================="
Write-Host " Merge-SPBackups"
Write-Host "============================================="
Write-Host ""

# ---------------------------------------------------------------------------
# Load files
# ---------------------------------------------------------------------------
Write-Host "Loading primary  : $PrimaryFile"
$pri = ConvertTo-WrappedBackup (ConvertFrom-JsonPreserveArrays (Get-Content -Path $PrimaryFile -Encoding UTF8 -Raw))
$priProjCountOrig = @($pri.data.project.ids).Count
$priTaskCountOrig = @($pri.data.task.ids).Count
$priTagCountOrig = (@($pri.data.tag.ids) | Where-Object { $_ -ne "TODAY" }).Count
$priRepeatCountOrig = @($pri.data.taskRepeatCfg.ids).Count
$priReminderCountOrig = if ($pri.data.reminders) { @($pri.data.reminders).Count } else { 0 }
$priSubtaskCountOrig = @($pri.data.task.ids | Where-Object {
        $t = $pri.data.task.entities[$_]
        $t -and $t.parentId
    }).Count

Write-Host "Loading secondary : $SecondaryFile"
$sec = ConvertTo-WrappedBackup (ConvertFrom-JsonPreserveArrays (Get-Content -Path $SecondaryFile -Encoding UTF8 -Raw))
$secTaskCountOrig = @($sec.data.task.ids).Count
$secProjCountOrig = @($sec.data.project.ids).Count
$secTagCountOrig = (@($sec.data.tag.ids) | Where-Object { $_ -ne "TODAY" }).Count
$secRepeatCountOrig = @($sec.data.taskRepeatCfg.ids).Count
$secReminderCountOrig = if ($sec.data.reminders) { @($sec.data.reminders).Count } else { 0 }
$secSubtaskCountOrig = @($sec.data.task.ids | Where-Object {
        $t = $sec.data.task.entities[$_]
        $t -and $t.parentId
    }).Count

$NowMs = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()

$priTaskIdSetOriginal = @{}
foreach ($id in @($pri.data.task.ids)) { $priTaskIdSetOriginal[$id] = $true }

# ID-remapping dictionaries: secondaryId â†’ primaryId
$projectMap = @{}
$tagMap = @{}
$taskMap = @{}
$repeatCfgMap = @{}

# Statistics
$stats = [ordered]@{
    ProjectsAdded = 0; ProjectsMerged = 0
    TagsAdded = 0; TagsMerged = 0
    RepeatCfgsAdded = 0; RepeatCfgsMerged = 0
    TasksKept = 0; TasksAdded = 0; TasksMerged = 0; TasksConflicted = 0
    SubtasksAdded = 0; SubtasksMerged = 0; SubtasksConflicted = 0
    RemindersAdded = 0
}

# ===========================================================================
# Phase 1 - Map PROJECTS
# ===========================================================================
Write-Host "`n--- Phase 1: Projects ---"

$priProjByTitle = @{}
foreach ($id in @($pri.data.project.ids)) {
    $p = $pri.data.project.entities[$id]
    if ($p -and $p.title) {
        $key = Normalize-Title $p.title
        if (-not $priProjByTitle.Contains($key)) { $priProjByTitle[$key] = $id }
    }
}

foreach ($id in @($sec.data.project.ids)) {
    $p = $sec.data.project.entities[$id]
    if (-not $p) { continue }

    # 1a. Same ID already in primary?
    if ($pri.data.project.entities.Contains($id)) {
        $projectMap[$id] = $id
        $stats.ProjectsMerged++
        Write-Host "  [MATCH-ID]    '$($p.title)' ($id)"
        continue
    }

    # 1b. Same title?
    $key = Normalize-Title $p.title
    if ($priProjByTitle.Contains($key)) {
        $primaryId = $priProjByTitle[$key]
        $projectMap[$id] = $primaryId
        $stats.ProjectsMerged++
        Write-Host "  [MATCH-TITLE] '$($p.title)' ($id -> $primaryId)"
        continue
    }

    # 1c. New project
    $projectMap[$id] = $id
    $pri.data.project.ids = @($pri.data.project.ids) + @($id)
    $pri.data.project.entities[$id] = $p
    $priProjByTitle[$key] = $id
    $stats.ProjectsAdded++
    Write-Host "  [ADD]         '$($p.title)' ($id)"
}

# ===========================================================================
# Phase 2 - Map TAGS
# ===========================================================================
Write-Host "`n--- Phase 2: Tags ---"

$priTagByTitle = @{}
foreach ($id in @($pri.data.tag.ids)) {
    $t = $pri.data.tag.entities[$id]
    if ($t -and $t.title) {
        $key = Normalize-Title $t.title
        if (-not $priTagByTitle.Contains($key)) { $priTagByTitle[$key] = $id }
    }
}

foreach ($id in @($sec.data.tag.ids)) {
    $t = $sec.data.tag.entities[$id]
    if (-not $t) { continue }

    # Well-known ID
    if ($id -eq "TODAY") {
        $tagMap["TODAY"] = "TODAY"
        Write-Host "  [MATCH-ID]    'Today'  (TODAY)"
        continue
    }

    # 2a. Same ID
    if ($pri.data.tag.entities.Contains($id)) {
        $tagMap[$id] = $id
        # Merge taskIds lists
        if ($t.taskIds) {
            foreach ($tid in $t.taskIds) {
                $pri.data.tag.entities[$id].taskIds = Safe-ArrayAdd $pri.data.tag.entities[$id].taskIds $tid
            }
        }
        $stats.TagsMerged++
        Write-Host "  [MATCH-ID]    '$($t.title)' ($id)"
        continue
    }

    # 2b. Same title
    $key = Normalize-Title $t.title
    if ($priTagByTitle.Contains($key)) {
        $primaryId = $priTagByTitle[$key]
        $tagMap[$id] = $primaryId
        $stats.TagsMerged++
        Write-Host "  [MATCH-TITLE] '$($t.title)' ($id -> $primaryId)"
        continue
    }

    # 2c. New tag - insert before TODAY so TODAY stays last
    $tagMap[$id] = $id
    $idsWithoutToday = @($pri.data.tag.ids | Where-Object { $_ -ne "TODAY" })
    $pri.data.tag.ids = $idsWithoutToday + @($id) + @("TODAY")
    $pri.data.tag.entities[$id] = $t
    $priTagByTitle[$key] = $id
    $stats.TagsAdded++
    Write-Host "  [ADD]         '$($t.title)' ($id)"
}

# ===========================================================================
# Phase 3 - Map REPEAT CONFIGS
# ===========================================================================
Write-Host "`n--- Phase 3: Repeat configs ---"

# Index primary repeat configs by (normalizedTitle, projectId)
$priRcByKey = @{}
foreach ($id in @($pri.data.taskRepeatCfg.ids)) {
    $rc = $pri.data.taskRepeatCfg.entities[$id]
    if (-not $rc) { continue }
    $key = "$(Normalize-Title $rc.title)|||$($rc.projectId)"
    if (-not $priRcByKey.Contains($key)) { $priRcByKey[$key] = $id }
}

foreach ($id in @($sec.data.taskRepeatCfg.ids)) {
    $rc = $sec.data.taskRepeatCfg.entities[$id]
    if (-not $rc) { continue }

    $mappedProjId = if ($projectMap.Contains($rc.projectId)) { $projectMap[$rc.projectId] } else { $rc.projectId }

    # 3a. Same ID
    if ($pri.data.taskRepeatCfg.entities.Contains($id)) {
        $repeatCfgMap[$id] = $id
        # Keep the one with newer lastTaskCreation
        $priRc = $pri.data.taskRepeatCfg.entities[$id]
        if ($rc.lastTaskCreation -gt $priRc.lastTaskCreation) {
            $rc.projectId = $mappedProjId
            # Remap tagIds
            $remapped = @()
            foreach ($tid in @($rc.tagIds)) {
                $remapped += if ($tagMap.Contains($tid)) { $tagMap[$tid] } else { $tid }
            }
            $rc.tagIds = $remapped
            $pri.data.taskRepeatCfg.entities[$id] = $rc
        }
        $stats.RepeatCfgsMerged++
        Write-Host "  [MATCH-ID]    '$($rc.title)' ($id)"
        continue
    }

    # 3b. Same title + mapped project
    $key = "$(Normalize-Title $rc.title)|||$mappedProjId"
    if ($priRcByKey.Contains($key)) {
        $primaryId = $priRcByKey[$key]
        $repeatCfgMap[$id] = $primaryId
        # Keep newer
        $priRc = $pri.data.taskRepeatCfg.entities[$primaryId]
        if ($rc.lastTaskCreation -gt $priRc.lastTaskCreation) {
            $rc.projectId = $mappedProjId
            $rc.id = $primaryId
            $remapped = @()
            foreach ($tid in @($rc.tagIds)) {
                $remapped += if ($tagMap.Contains($tid)) { $tagMap[$tid] } else { $tid }
            }
            $rc.tagIds = $remapped
            $pri.data.taskRepeatCfg.entities[$primaryId] = $rc
        }
        $stats.RepeatCfgsMerged++
        Write-Host "  [MATCH-TITLE] '$($rc.title)' ($id -> $primaryId)"
        continue
    }

    # 3c. New repeat config
    $repeatCfgMap[$id] = $id
    $rc.projectId = $mappedProjId
    $remapped = @()
    foreach ($tid in @($rc.tagIds)) {
        $remapped += if ($tagMap.Contains($tid)) { $tagMap[$tid] } else { $tid }
    }
    $rc.tagIds = $remapped
    $pri.data.taskRepeatCfg.ids = @($pri.data.taskRepeatCfg.ids) + @($id)
    $pri.data.taskRepeatCfg.entities[$id] = $rc
    $stats.RepeatCfgsAdded++
    Write-Host "  [ADD]         '$($rc.title)' ($id)"
}

# ===========================================================================
# Phase 4 - Map and merge TASKS
# ===========================================================================
Write-Host "`n--- Phase 4: Tasks ---"

# Index primary tasks by (normalizedTitle, projectId) â†’ list of IDs
$priTaskIndex = @{}
foreach ($id in @($pri.data.task.ids)) {
    $t = $pri.data.task.entities[$id]
    if (-not $t) { continue }
    $key = "$(Normalize-Title $t.title)|||$($t.projectId)"
    if (-not $priTaskIndex.Contains($key)) {
        $priTaskIndex[$key] = [System.Collections.Generic.List[string]]::new()
    }
    $priTaskIndex[$key].Add($id)
}

$priTaskIdSet = @{}
foreach ($id in @($pri.data.task.ids)) { $priTaskIdSet[$id] = $true }

# Separate non-subtasks and subtasks (process parents first)
$secParents = @()
$secChildren = @()
foreach ($id in @($sec.data.task.ids)) {
    $t = $sec.data.task.entities[$id]
    if (-not $t) { continue }
    if ($t.Contains("parentId") -and $t.parentId) {
        $secChildren += @{ id = $id; task = $t }
    }
    else {
        $secParents += @{ id = $id; task = $t }
    }
}

# ---- Helper: remap IDs inside a task and register it into primary ----
function Remap-TaskIds {
    param($task, [string]$taskId)

    # projectId
    if ($projectMap.Contains($task.projectId)) {
        $task.projectId = $projectMap[$task.projectId]
    }

    # tagIds
    if ($task.tagIds) {
        $remapped = @()
        foreach ($tid in @($task.tagIds)) {
            $mapped = if ($tagMap.Contains($tid)) { $tagMap[$tid] } else { $tid }
            $remapped += $mapped
            # Update tag entity's taskIds
            if ($pri.data.tag.entities.Contains($mapped)) {
                $pri.data.tag.entities[$mapped].taskIds = Safe-ArrayAdd $pri.data.tag.entities[$mapped].taskIds $taskId
            }
        }
        $task.tagIds = $remapped
    }

    # repeatCfgId
    if ($task.Contains("repeatCfgId") -and $task.repeatCfgId) {
        if ($repeatCfgMap.Contains($task.repeatCfgId)) {
            $task.repeatCfgId = $repeatCfgMap[$task.repeatCfgId]
        }
    }

    # parentId (for subtasks)
    if ($task.Contains("parentId") -and $task.parentId) {
        if ($taskMap.Contains($task.parentId)) {
            $task.parentId = $taskMap[$task.parentId]
        }
    }
}

# ---- Pre-assign a final id to every secondary task ------------------------
# Conflict policy: nothing in the primary is overwritten. Every secondary task
# becomes its own entry. A secondary id is kept as-is unless it collides with an
# existing primary task id, in which case a fresh unique id is minted so both
# tasks can coexist. This map also rewrites parentId / subTaskIds references
# between secondary tasks.
foreach ($id in @($pri.data.task.ids)) { $script:usedTaskIds[$id] = $true }
foreach ($id in @($sec.data.task.ids)) { $script:usedTaskIds[$id] = $true }
foreach ($id in @($sec.data.task.ids)) {
    if ($priTaskIdSetOriginal.Contains($id)) {
        $taskMap[$id] = New-UniqueTaskId   # id collision -> new id for the secondary copy
    }
    else {
        $taskMap[$id] = $id
    }
}

# ---- Process parent tasks ----
foreach ($item in $secParents) {
    $secId   = $item.id
    $secTask = $item.task
    $finalId = $taskMap[$secId]

    $mappedProjId = if ($projectMap.Contains($secTask.projectId)) { $projectMap[$secTask.projectId] } else { $secTask.projectId }

    # ---- Detect conflict against the ORIGINAL primary tasks ----
    #   conflict = same id, OR same title in the same mapped project with a
    #   matching instance anchor (the cases that previously caused a merge).
    $isConflict = $priTaskIdSetOriginal.Contains($secId)
    if (-not $isConflict) {
        $key = "$(Normalize-Title $secTask.title)|||$mappedProjId"
        if ($priTaskIndex.Contains($key)) {
            $mappedRepeatCfgId = $null
            if ($secTask.Contains("repeatCfgId") -and $secTask.repeatCfgId) {
                $mappedRepeatCfgId = if ($repeatCfgMap.Contains($secTask.repeatCfgId)) { $repeatCfgMap[$secTask.repeatCfgId] } else { $secTask.repeatCfgId }
            }
            foreach ($cId in $priTaskIndex[$key]) {
                if (-not $priTaskIdSetOriginal.Contains($cId)) { continue }
                if (Is-SameTaskInstance $pri.data.task.entities[$cId] $secTask $mappedRepeatCfgId) {
                    $isConflict = $true; break
                }
            }
        }
    }

    if ($isConflict) {
        $secTask.title = Add-ConflictPrefix $secTask.title
        $stats.TasksConflicted++
    }

    # ---- Add the secondary task as its own entry (primary untouched) ----
    $taskMap[$secId] = $finalId
    $secTask.id      = $finalId
    Remap-TaskIds $secTask $finalId
    if ($secTask.Contains("subTaskIds") -and $secTask.subTaskIds) {
        $remappedSubs = @()
        foreach ($st in @($secTask.subTaskIds)) {
            $remappedSubs += if ($taskMap.Contains($st)) { $taskMap[$st] } else { $st }
        }
        $secTask.subTaskIds = $remappedSubs
    }

    $pri.data.task.ids = @($pri.data.task.ids) + @($finalId)
    $pri.data.task.entities[$finalId] = $secTask
    $priTaskIdSet[$finalId] = $true

    # Add to project's taskIds
    if ($pri.data.project.entities.Contains($secTask.projectId)) {
        $pri.data.project.entities[$secTask.projectId].taskIds = `
            Safe-ArrayAdd $pri.data.project.entities[$secTask.projectId].taskIds $finalId
    }

    # Update index
    $ikey = "$(Normalize-Title $secTask.title)|||$($secTask.projectId)"
    if (-not $priTaskIndex.Contains($ikey)) {
        $priTaskIndex[$ikey] = [System.Collections.Generic.List[string]]::new()
    }
    $priTaskIndex[$ikey].Add($finalId)

    $stats.TasksAdded++
    if ($isConflict) {
        Write-Host "  [CONFLICT]    '$($secTask.title)' (project=$($secTask.projectId))"
    }
    else {
        Write-Host "  [ADD]         '$($secTask.title)' (project=$($secTask.projectId))"
    }
}

# ---- Process subtasks ----
Write-Host ""
foreach ($item in $secChildren) {
    $secId   = $item.id
    $secTask = $item.task
    $finalId = $taskMap[$secId]

    $mappedProjId   = if ($projectMap.Contains($secTask.projectId)) { $projectMap[$secTask.projectId] } else { $secTask.projectId }
    $mappedParentId = if ($taskMap.Contains($secTask.parentId))     { $taskMap[$secTask.parentId] }     else { $secTask.parentId }

    # ---- Detect conflict against the ORIGINAL primary tasks ----
    $isConflict = $priTaskIdSetOriginal.Contains($secId)
    if (-not $isConflict) {
        $key = "$(Normalize-Title $secTask.title)|||$mappedProjId"
        if ($priTaskIndex.Contains($key)) {
            $mappedRepeatCfgId = $null
            if ($secTask.Contains("repeatCfgId") -and $secTask.repeatCfgId) {
                $mappedRepeatCfgId = if ($repeatCfgMap.Contains($secTask.repeatCfgId)) { $repeatCfgMap[$secTask.repeatCfgId] } else { $secTask.repeatCfgId }
            }
            foreach ($cId in $priTaskIndex[$key]) {
                if (-not $priTaskIdSetOriginal.Contains($cId)) { continue }
                $c = $pri.data.task.entities[$cId]
                if ($c.Contains("parentId") -and $c.parentId -eq $mappedParentId) {
                    if (Is-SameTaskInstance $c $secTask $mappedRepeatCfgId) {
                        $isConflict = $true; break
                    }
                }
            }
        }
    }

    if ($isConflict) {
        $secTask.title = Add-ConflictPrefix $secTask.title
        $stats.SubtasksConflicted++
    }

    # ---- Add the secondary subtask as its own entry (primary untouched) ----
    $taskMap[$secId]   = $finalId
    $secTask.id        = $finalId
    $secTask.projectId = $mappedProjId
    $secTask.parentId  = $mappedParentId
    Remap-TaskIds $secTask $finalId
    if ($secTask.Contains("subTaskIds") -and $secTask.subTaskIds) {
        $remappedSubs = @()
        foreach ($st in @($secTask.subTaskIds)) {
            $remappedSubs += if ($taskMap.Contains($st)) { $taskMap[$st] } else { $st }
        }
        $secTask.subTaskIds = $remappedSubs
    }

    $pri.data.task.ids = @($pri.data.task.ids) + @($finalId)
    $pri.data.task.entities[$finalId] = $secTask
    $priTaskIdSet[$finalId] = $true

    # Attach to the (mapped) parent's subTaskIds so the copy stays nested
    if ($pri.data.task.entities.Contains($mappedParentId)) {
        $pri.data.task.entities[$mappedParentId].subTaskIds = `
            Safe-ArrayAdd $pri.data.task.entities[$mappedParentId].subTaskIds $finalId
    }

    # Update index
    $ikey = "$(Normalize-Title $secTask.title)|||$mappedProjId"
    if (-not $priTaskIndex.Contains($ikey)) {
        $priTaskIndex[$ikey] = [System.Collections.Generic.List[string]]::new()
    }
    $priTaskIndex[$ikey].Add($finalId)

    $stats.SubtasksAdded++
    if ($isConflict) {
        Write-Host "  [CONFLICT-SUB] '$($secTask.title)' (parent=$mappedParentId)"
    }
    else {
        Write-Host "  [ADD-SUB]     '$($secTask.title)' (parent=$mappedParentId)"
    }
}

# ===========================================================================
# Phase 5 - Fixup: remap tag.taskIds using taskMap
# ===========================================================================
Write-Host "`n--- Phase 5: Fixup tag.taskIds ---"

foreach ($tagId in @($pri.data.tag.entities.Keys)) {
    $tag = $pri.data.tag.entities[$tagId]
    if (-not $tag.taskIds) { continue }
    $fixed = @()
    foreach ($tid in @($tag.taskIds)) {
        $mapped = if ($taskMap.Contains($tid)) { $taskMap[$tid] } else { $tid }
        if ($mapped -notin $fixed) { $fixed += $mapped }
    }
    $tag.taskIds = $fixed
}

# ===========================================================================
# Phase 6 - Merge REMINDERS
# ===========================================================================
Write-Host "`n--- Phase 6: Reminders ---"

if (-not $pri.data.Contains("reminders") -or $null -eq $pri.data.reminders) {
    $pri.data["reminders"] = @()
}

# Index primary reminders by relatedId
$priReminderByRelated = @{}
foreach ($r in @($pri.data.reminders)) {
    if ($r -and $r.Contains("relatedId") -and $r.relatedId) {
        $priReminderByRelated[$r.relatedId] = $r
    }
}

if ($sec.data.Contains("reminders") -and $sec.data.reminders) {
    foreach ($r in @($sec.data.reminders)) {
        if (-not $r -or -not $r.Contains("relatedId")) { continue }

        $mappedRelated = if ($taskMap.Contains($r.relatedId)) { $taskMap[$r.relatedId] } else { $r.relatedId }

        if ($priReminderByRelated.Contains($mappedRelated)) {
            # Already have a reminder for this task - keep primary's
            Write-Host "  [SKIP]        Reminder for task $mappedRelated"
            continue
        }

        # Add new reminder with remapped relatedId
        $r.relatedId = $mappedRelated
        $pri.data.reminders = @($pri.data.reminders) + @($r)
        $priReminderByRelated[$mappedRelated] = $r
        $stats.RemindersAdded++
        Write-Host "  [ADD]         Reminder '$($r.title)' -> task $mappedRelated"
    }
}

# ===========================================================================
# Phase 7 - Merge MENU TREE
# ===========================================================================
Write-Host "`n--- Phase 7: Menu tree ---"

# Collect all project/tag IDs already present in the tree
function Get-TreeIds {
    param([array]$tree)
    $ids = @{}
    foreach ($node in @($tree)) {
        if (-not $node) { continue }
        if ($node.Contains("id") -and ($node.Contains("k"))) {
            if ($node.k -eq "p" -or $node.k -eq "t") { $ids[$node.id] = $true }
        }
        if ($node.Contains("children") -and $node.children) {
            $childIds = Get-TreeIds $node.children
            foreach ($k in $childIds.Keys) { $ids[$k] = $true }
        }
    }
    return $ids
}

# --- Project tree ---
$existingProjTreeIds = Get-TreeIds $pri.data.menuTree.projectTree

$newProjNodes = @()
foreach ($id in @($sec.data.project.ids)) {
    $mappedId = if ($projectMap.Contains($id)) { $projectMap[$id] } else { $id }
    if ($mappedId -eq "INBOX_PROJECT") { continue }
    if (-not $existingProjTreeIds.Contains($mappedId)) {
        $newProjNodes += [ordered]@{ k = "p"; id = $mappedId }
        $existingProjTreeIds[$mappedId] = $true
        Write-Host "  [TREE-PROJ]   Added $mappedId"
    }
}
if ($newProjNodes.Count -gt 0) {
    $pri.data.menuTree.projectTree = @($pri.data.menuTree.projectTree) + @(
        [ordered]@{
            k          = "f"
            id         = [guid]::NewGuid().ToString()
            name       = "Merged-Import"
            isExpanded = $true
            children   = $newProjNodes
        }
    )
}

# --- Tag tree ---
$existingTagTreeIds = Get-TreeIds $pri.data.menuTree.tagTree

$newTagNodes = @()
foreach ($id in @($sec.data.tag.ids)) {
    if ($id -eq "TODAY") { continue }
    $mappedId = if ($tagMap.Contains($id)) { $tagMap[$id] } else { $id }
    if (-not $existingTagTreeIds.Contains($mappedId)) {
        $newTagNodes += [ordered]@{ k = "t"; id = $mappedId }
        $existingTagTreeIds[$mappedId] = $true
        Write-Host "  [TREE-TAG]    Added $mappedId"
    }
}
if ($newTagNodes.Count -gt 0) {
    # Insert folder before the final TODAY entry
    $todayNode = $pri.data.menuTree.tagTree | Where-Object { $_.id -eq "TODAY" } | Select-Object -First 1
    $otherNodes = @($pri.data.menuTree.tagTree | Where-Object { $_.id -ne "TODAY" })
    $newFolder = [ordered]@{
        k          = "f"
        id         = [guid]::NewGuid().ToString()
        name       = "Merged-Import Tags"
        isExpanded = $true
        children   = $newTagNodes
    }
    $pri.data.menuTree.tagTree = $otherNodes + @($newFolder)
    if ($todayNode) { $pri.data.menuTree.tagTree += $todayNode }
}

# ===========================================================================
# Phase 8 - Merge other collections
# ===========================================================================
Write-Host "`n--- Phase 8: Other data ---"

# --- Notes ---
if ($sec.data.Contains("note") -and $sec.data.note.ids) {
    if (-not $pri.data.Contains("note")) {
        $pri.data["note"] = [ordered]@{ ids = @(); entities = [ordered]@{}; todayOrder = @() }
    }
    foreach ($id in @($sec.data.note.ids)) {
        if (-not $pri.data.note.entities.Contains($id)) {
            $pri.data.note.ids = @($pri.data.note.ids) + @($id)
            $pri.data.note.entities[$id] = $sec.data.note.entities[$id]
            Write-Host "  [ADD-NOTE]    $id"
        }
    }
}

# --- Simple counters ---
if ($sec.data.Contains("simpleCounter") -and $sec.data.simpleCounter.ids) {
    if (-not $pri.data.Contains("simpleCounter")) {
        $pri.data["simpleCounter"] = [ordered]@{ ids = @(); entities = [ordered]@{} }
    }
    foreach ($id in @($sec.data.simpleCounter.ids)) {
        if (-not $pri.data.simpleCounter.entities.Contains($id)) {
            $pri.data.simpleCounter.ids = @($pri.data.simpleCounter.ids) + @($id)
            $pri.data.simpleCounter.entities[$id] = $sec.data.simpleCounter.entities[$id]
            Write-Host "  [ADD-COUNTER] $id"
        }
    }
}

# --- Metrics ---
if ($sec.data.Contains("metric") -and $sec.data.metric.ids) {
    if (-not $pri.data.Contains("metric")) {
        $pri.data["metric"] = [ordered]@{ ids = @(); entities = [ordered]@{} }
    }
    foreach ($id in @($sec.data.metric.ids)) {
        if (-not $pri.data.metric.entities.Contains($id)) {
            $pri.data.metric.ids = @($pri.data.metric.ids) + @($id)
            $pri.data.metric.entities[$id] = $sec.data.metric.entities[$id]
            Write-Host "  [ADD-METRIC]  $id"
        }
    }
}

# --- Issue providers ---
if ($sec.data.Contains("issueProvider") -and $sec.data.issueProvider.ids) {
    if (-not $pri.data.Contains("issueProvider")) {
        $pri.data["issueProvider"] = [ordered]@{ ids = @(); entities = [ordered]@{} }
    }
    foreach ($id in @($sec.data.issueProvider.ids)) {
        if (-not $pri.data.issueProvider.entities.Contains($id)) {
            $pri.data.issueProvider.ids = @($pri.data.issueProvider.ids) + @($id)
            $pri.data.issueProvider.entities[$id] = $sec.data.issueProvider.entities[$id]
            Write-Host "  [ADD-ISSUE]   $id"
        }
    }
}

# --- Boards ---
if ($sec.data.Contains("boards") -and $sec.data.boards.boardCfgs) {
    if (-not $pri.data.Contains("boards") -or -not $pri.data.boards.boardCfgs) {
        $pri.data.boards = [ordered]@{ boardCfgs = @() }
    }
    foreach ($cfg in @($sec.data.boards.boardCfgs)) {
        if (-not $cfg) { continue }
        $found = $false
        foreach ($existing in @($pri.data.boards.boardCfgs)) {
            if ($existing -and $existing.Contains("id") -and $cfg.Contains("id") -and $existing.id -eq $cfg.id) {
                $found = $true; break
            }
        }
        if (-not $found) {
            $pri.data.boards.boardCfgs = @($pri.data.boards.boardCfgs) + @($cfg)
            Write-Host "  [ADD-BOARD]   $($cfg.id)"
        }
    }
}

# --- Planner days ---
if ($sec.data.Contains("planner") -and $sec.data.planner.days) {
    foreach ($day in @($sec.data.planner.days.Keys)) {
        if (-not $pri.data.planner.days.Contains($day)) {
            $pri.data.planner.days[$day] = $sec.data.planner.days[$day]
            Write-Host "  [ADD-PLANNER] $day"
        }
    }
}

# --- Time tracking ---
if ($sec.data.Contains("timeTracking") -and $sec.data.timeTracking) {
    foreach ($scope in @("project", "tag")) {
        if ($sec.data.timeTracking.Contains($scope) -and $sec.data.timeTracking[$scope]) {
            foreach ($entityId in @($sec.data.timeTracking[$scope].Keys)) {
                $mappedId = $entityId
                if ($scope -eq "project" -and $projectMap.Contains($entityId)) { $mappedId = $projectMap[$entityId] }
                if ($scope -eq "tag" -and $tagMap.Contains($entityId)) { $mappedId = $tagMap[$entityId] }

                if (-not $pri.data.timeTracking[$scope].Contains($mappedId)) {
                    $pri.data.timeTracking[$scope][$mappedId] = $sec.data.timeTracking[$scope][$entityId]
                }
                else {
                    # Merge day-level entries
                    $secTt = $sec.data.timeTracking[$scope][$entityId]
                    $priTt = $pri.data.timeTracking[$scope][$mappedId]
                    if ($secTt -is [hashtable] -or $secTt -is [ordered]) {
                        foreach ($dk in @($secTt.Keys)) {
                            if (-not $priTt.Contains($dk)) {
                                $priTt[$dk] = $secTt[$dk]
                            }
                        }
                    }
                }
            }
        }
    }
}

# --- Archives (Young / Old) ---
foreach ($archName in @("archiveYoung", "archiveOld")) {
    if (-not $sec.data.Contains($archName) -or -not $sec.data[$archName]) { continue }
    if (-not $pri.data.Contains($archName)) {
        $pri.data[$archName] = [ordered]@{
            task                  = [ordered]@{ ids = @(); entities = [ordered]@{} }
            timeTracking          = [ordered]@{ project = [ordered]@{}; tag = [ordered]@{} }
            lastTimeTrackingFlush = 0
        }
    }

    $secArch = $sec.data[$archName]
    $priArch = $pri.data[$archName]

    if ($secArch.Contains("task") -and $secArch.task.ids) {
        foreach ($id in @($secArch.task.ids)) {
            if (-not $priArch.task.entities.Contains($id)) {
                $archTask = $secArch.task.entities[$id]
                # Remap IDs in archived task
                if ($archTask) {
                    if ($projectMap.Contains($archTask.projectId)) {
                        $archTask.projectId = $projectMap[$archTask.projectId]
                    }
                    if ($archTask.Contains("tagIds") -and $archTask.tagIds) {
                        $remappedArchTags = @()
                        foreach ($tid in @($archTask.tagIds)) {
                            $remappedArchTags += if ($tagMap.Contains($tid)) { $tagMap[$tid] } else { $tid }
                        }
                        $archTask.tagIds = $remappedArchTags
                    }
                    if ($archTask.Contains("repeatCfgId") -and $archTask.repeatCfgId) {
                        if ($repeatCfgMap.Contains($archTask.repeatCfgId)) {
                            $archTask.repeatCfgId = $repeatCfgMap[$archTask.repeatCfgId]
                        }
                    }
                }
                $priArch.task.ids = @($priArch.task.ids) + @($id)
                $priArch.task.entities[$id] = $archTask
                Write-Host "  [ADD-ARCH]    $archName task '$($archTask.title)'"
            }
        }
    }
}

# --- Plugin data ---
foreach ($field in @("pluginMetadata", "pluginUserData")) {
    if ($sec.data.Contains($field) -and $sec.data[$field]) {
        if (-not $pri.data.Contains($field)) { $pri.data[$field] = @() }
        foreach ($item in @($sec.data[$field])) {
            if (-not $item) { continue }
            $pri.data[$field] = @($pri.data[$field]) + @($item)
        }
    }
}

# --- globalConfig: keep primary's (user's actual settings) ---
# No merge needed - primary's config is authoritative.

# ===========================================================================
# Phase 9 – Finalize & Output
# ===========================================================================
# Array sanitization is no longer needed here: ConvertFrom-JsonPreserveArrays
# preserves every JSON array as a proper @() at parse time, and all merge
# operations build arrays with @(...) + @(...) syntax.

# Update top-level timestamps
$pri.timestamp = $NowMs
$pri.lastUpdate = $NowMs
if (-not $pri.Contains("crossModelVersion")) { $pri["crossModelVersion"] = 4.5 }

# Serialize
# --------------------------------------------------------------------------
# Safety net: walk the entire tree and normalise any collection-wrapper
# objects that PowerShell may have created (e.g. ArrayList, PSObject wrapper)
# back to plain [object[]] arrays so that ConvertTo-Json serialises them as
# JSON arrays rather than {"value":[],"Count":0} objects.
function Normalize-Collections {
    param($obj)
    if ($null -eq $obj) { return $null }

    if ($obj -is [System.Collections.IDictionary]) {
        $keys = @($obj.Keys)
        foreach ($k in $keys) {
            $obj[$k] = Normalize-Collections $obj[$k]
        }
        return $obj
    }

    if ($obj -is [string]) { return $obj }

    if ($obj -is [System.Collections.IEnumerable]) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $obj) {
            $list.Add((Normalize-Collections $item))
        }
        # IMPORTANT: emit the array WITHOUT pipeline-unrolling.
        # A bare `return <array>` is unrolled by PowerShell when the caller
        # captures it: an empty array collapses to $null and a single-element
        # array collapses to the bare element. That is exactly what corrupts
        # the array-valued fields (tagIds / subTaskIds / attachments / ...),
        # so the very function meant to preserve arrays was destroying them.
        # -NoEnumerate forces the array onto the stream as one object.
        Write-Output -NoEnumerate ([object[]]$list.ToArray())
        return
    }

    return $obj
}

# ---------------------------------------------------------------------------
# Schema-aware array coercion (authoritative safety net)
# ---------------------------------------------------------------------------
# Super Productivity's import validator requires certain fields to be JSON
# arrays. PowerShell can unwrap a single-element array into a bare scalar, turn
# an empty array into $null, or (under Windows PowerShell 5.1) serialise an
# empty collection as {}. None of those can be recovered generically, because a
# bare string or $null carries no hint that it was meant to be an array. The
# fix has to be schema-aware: for every known array-valued field on every known
# entity, force the value to a real [object[]] using DIRECT ASSIGNMENT (which,
# unlike a function return, never triggers pipeline unrolling).
#
# Canonical array-valued fields per entity type, derived from the SP data model
# (and verified against the real backup + export files).
$script:EntityArrayFields = @{
    task          = @("subTaskIds", "tagIds", "attachments")
    project       = @("taskIds", "backlogTaskIds", "noteIds")
    tag           = @("taskIds")
    taskRepeatCfg = @("tagIds", "deletedInstanceDates")
    issueProvider = @("defaultTagIds")
    metric        = @("focusSessions", "reflections")
    # Archived tasks share the task schema plus a legacy "subTasks" array.
    archiveTask   = @("subTaskIds", "tagIds", "attachments", "subTasks")
}

# Coerce a single value to a genuine [object[]] without ever unrolling.
# Returns via the caller's direct assignment, so this is a *filter-free* helper:
# callers must use it only as  $x = ... then re-wrap, OR use Set-ArrayField.
function Set-ArrayField {
    param($entity, [string]$field)
    if ($null -eq $entity -or -not ($entity -is [System.Collections.IDictionary])) { return }
    $cur = if ($entity.Contains($field)) { $entity[$field] } else { $null }

    if ($null -eq $cur) {
        $entity[$field] = [object[]]@()
    }
    elseif ($cur -is [string]) {
        # single id that was unwrapped from a one-element array
        $entity[$field] = [object[]]@($cur)
    }
    elseif ($cur -is [System.Collections.IDictionary]) {
        if ($cur.Count -eq 0) {
            # empty array that PowerShell rendered as an empty object {}
            $entity[$field] = [object[]]@()
        }
        else {
            # a single real element (e.g. one attachment) unwrapped from its array
            $entity[$field] = [object[]]@($cur)
        }
    }
    elseif ($cur -is [System.Collections.IEnumerable]) {
        # already a collection (incl. a proper array) -> normalise to [object[]]
        $entity[$field] = [object[]]@($cur)
    }
    else {
        # scalar number / bool that was a one-element array
        $entity[$field] = [object[]]@($cur)
    }
}

# Walk an entities dictionary and coerce every listed array field in place.
function Coerce-EntityCollection {
    param($collection, [string[]]$fields)
    if ($null -eq $collection -or -not ($collection -is [System.Collections.IDictionary])) { return }
    $entities = if ($collection.Contains("entities")) { $collection["entities"] } else { $null }
    if ($null -eq $entities -or -not ($entities -is [System.Collections.IDictionary])) { return }
    foreach ($eid in @($entities.Keys)) {
        $e = $entities[$eid]
        foreach ($f in $fields) { Set-ArrayField $e $f }
    }
}

# Apply the schema-aware coercion across the whole backup.
function Coerce-AllArrayFields {
    param($backup)
    if ($null -eq $backup -or -not $backup.Contains("data")) { return }
    $data = $backup["data"]

    Coerce-EntityCollection $data["task"]          $script:EntityArrayFields.task
    Coerce-EntityCollection $data["project"]       $script:EntityArrayFields.project
    Coerce-EntityCollection $data["tag"]           $script:EntityArrayFields.tag
    Coerce-EntityCollection $data["taskRepeatCfg"] $script:EntityArrayFields.taskRepeatCfg
    Coerce-EntityCollection $data["issueProvider"] $script:EntityArrayFields.issueProvider
    Coerce-EntityCollection $data["metric"]        $script:EntityArrayFields.metric

    # Archived task collections (archiveYoung / archiveOld).
    foreach ($archName in @("archiveYoung", "archiveOld")) {
        if ($data.Contains($archName) -and $data[$archName] -and $data[$archName].Contains("task")) {
            Coerce-EntityCollection $data[$archName]["task"] $script:EntityArrayFields.archiveTask
        }
    }

    # Top-level reminders must be an array as well.
    if (-not $data.Contains("reminders") -or $null -eq $data["reminders"]) {
        $data["reminders"] = [object[]]@()
    }
    elseif (-not ($data["reminders"] -is [System.Collections.IEnumerable]) -or ($data["reminders"] -is [System.Collections.IDictionary])) {
        if (($data["reminders"] -is [System.Collections.IDictionary]) -and $data["reminders"].Count -eq 0) {
            $data["reminders"] = [object[]]@()
        }
        else {
            $data["reminders"] = [object[]]@($data["reminders"])
        }
    }
    else {
        $data["reminders"] = [object[]]@($data["reminders"])
    }
}

$pri = Normalize-Collections $pri
# Authoritative pass: guarantee every schema-required array field is a real
# array regardless of how an upstream step may have unwrapped it.
Coerce-AllArrayFields $pri

# ---------------------------------------------------------------------------
# Deterministic JSON serializer (version-independent)
# ---------------------------------------------------------------------------
# We do NOT use ConvertTo-Json. Under Windows PowerShell 5.1 its serializer
# reflects certain collection types (e.g. the arrays produced by `@(..) + @(..)`
# and the .ids lists) into wrapper objects like {"value":[...],"Count":N},
# which Super Productivity rejects ("task.ids expected Array<string>",
# "note.todayOrder.filter is not a function"). This hand-rolled writer depends
# only on type checks and a StringBuilder, so it emits identical, correct JSON
# on both PowerShell 5.1 and 7+: ordered dictionaries become JSON objects and
# every non-string IEnumerable becomes a JSON array (empty -> [], single -> [x]).
function ConvertTo-SpJsonString {
    param([string]$s)
    if ($null -eq $s) { return '""' }
    $sb = [System.Text.StringBuilder]::new($s.Length + 2)
    [void]$sb.Append('"')
    foreach ($ch in $s.ToCharArray()) {
        switch ($ch) {
            '"'  { [void]$sb.Append('\"') }
            '\'  { [void]$sb.Append('\\') }
            "`b" { [void]$sb.Append('\b') }
            "`f" { [void]$sb.Append('\f') }
            "`n" { [void]$sb.Append('\n') }
            "`r" { [void]$sb.Append('\r') }
            "`t" { [void]$sb.Append('\t') }
            default {
                if ([int]$ch -lt 0x20) {
                    [void]$sb.Append('\u')
                    [void]$sb.Append(([int]$ch).ToString('x4'))
                }
                else {
                    [void]$sb.Append($ch)
                }
            }
        }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Write-SpJsonValue {
    param($val, [System.Text.StringBuilder]$sb)

    if ($null -eq $val) { [void]$sb.Append('null'); return }

    if ($val -is [string]) { [void]$sb.Append((ConvertTo-SpJsonString $val)); return }
    if ($val -is [bool])   { [void]$sb.Append($(if ($val) { 'true' } else { 'false' })); return }

    if ($val -is [int] -or $val -is [long] -or $val -is [int64] -or $val -is [int32] -or
        $val -is [int16] -or $val -is [byte] -or $val -is [sbyte] -or $val -is [uint32] -or $val -is [uint64]) {
        [void]$sb.Append(([System.IConvertible]$val).ToString([System.Globalization.CultureInfo]::InvariantCulture)); return
    }
    if ($val -is [double] -or $val -is [single] -or $val -is [decimal]) {
        $d = [double]$val
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { [void]$sb.Append('null'); return }
        [void]$sb.Append($d.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)); return
    }

    if ($val -is [System.Collections.IDictionary]) {
        [void]$sb.Append('{')
        $first = $true
        foreach ($k in $val.Keys) {
            if (-not $first) { [void]$sb.Append(',') }
            $first = $false
            [void]$sb.Append((ConvertTo-SpJsonString ([string]$k)))
            [void]$sb.Append(':')
            Write-SpJsonValue $val[$k] $sb
        }
        [void]$sb.Append('}')
        return
    }

    # Any non-string enumerable is a JSON array. This is the key line: it makes
    # empty collections serialise as [] and single-element ones as [x], with no
    # {value,Count} wrapper, regardless of the concrete collection type.
    if ($val -is [System.Collections.IEnumerable]) {
        [void]$sb.Append('[')
        $first = $true
        foreach ($item in $val) {
            if (-not $first) { [void]$sb.Append(',') }
            $first = $false
            Write-SpJsonValue $item $sb
        }
        [void]$sb.Append(']')
        return
    }

    # Fallback: stringify any unknown scalar type.
    [void]$sb.Append((ConvertTo-SpJsonString ([string]$val)))
}

$__sb = [System.Text.StringBuilder]::new(16 * 1024 * 1024)
Write-SpJsonValue $pri $__sb
$json = $__sb.ToString()

$resolvedOutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
$outDir = Split-Path $resolvedOutputFile -Parent
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

[System.IO.File]::WriteAllText(
    $resolvedOutputFile,
    $json,
    (New-Object System.Text.UTF8Encoding($false))   # UTF-8 without BOM (SP import rejects a BOM)
)

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$totalTasks = @($pri.data.task.ids).Count
$totalProjects = @($pri.data.project.ids).Count
$totalTags = @($pri.data.tag.ids).Count - 1  # exclude TODAY
$totalRepeats = @($pri.data.taskRepeatCfg.ids).Count
$totalReminders = @($pri.data.reminders).Count

Write-Host "`n============================================="
Write-Host " Merge complete"
Write-Host "============================================="
Write-Host ""
Write-Host "  Source stats:"
$fmtHeader = "    {0,-12} {1,8} {2,10}"
Write-Host ($fmtHeader -f "", "Primary", "Secondary")
Write-Host ($fmtHeader -f "Projects:",  $priProjCountOrig, $secProjCountOrig)
Write-Host ($fmtHeader -f "Tasks:",     $priTaskCountOrig, $secTaskCountOrig)
Write-Host ($fmtHeader -f "Subtasks:",  $priSubtaskCountOrig, $secSubtaskCountOrig)
Write-Host ($fmtHeader -f "Tags:",      $priTagCountOrig, $secTagCountOrig)
Write-Host ($fmtHeader -f "Repeats:",   $priRepeatCountOrig, $secRepeatCountOrig)
Write-Host ($fmtHeader -f "Reminders:", $priReminderCountOrig, $secReminderCountOrig)
Write-Host ""
Write-Host "  Merge actions:"
Write-Host "    Projects  : $($stats.ProjectsMerged) matched, $($stats.ProjectsAdded) added"
Write-Host "    Tags      : $($stats.TagsMerged) matched, $($stats.TagsAdded) added"
Write-Host "    Tasks     : $($stats.TasksAdded) added ($($stats.TasksConflicted) flagged [CONFLICT])"
Write-Host "    Subtasks  : $($stats.SubtasksAdded) added ($($stats.SubtasksConflicted) flagged [CONFLICT])"
Write-Host "    RepeatCfgs: $($stats.RepeatCfgsMerged) matched, $($stats.RepeatCfgsAdded) added"
Write-Host "    Reminders : $($stats.RemindersAdded) added"
Write-Host ""
Write-Host "  Output totals:"
Write-Host "    Projects  : $totalProjects"
Write-Host "    Tasks     : $totalTasks (incl. subtasks)"
Write-Host "    Tags      : $totalTags (+ TODAY)"
Write-Host "    RepeatCfgs: $totalRepeats"
Write-Host "    Reminders : $totalReminders"
Write-Host "    Output    : $resolvedOutputFile"
