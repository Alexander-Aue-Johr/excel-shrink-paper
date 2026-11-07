param([Parameter(Mandatory = $true)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$directory = $config.directory
$utf8 = New-Object Text.UTF8Encoding($false)
$stage = 'startup'
$timings = [ordered]@{}
$stageTimer = [Diagnostics.Stopwatch]::StartNew()
$owned = $null
$excel = $null
$native = $null
$workbook = $null
$rebuilt = $null
$copyObjectsSetting = $null
$clipboardMutex = $null
$clipboardOwned = $false
$result = [ordered]@{ status = 'review'; stage = 'startup'; timings = $timings; memory = @() }

function Write-Json([string]$path, $value) {
    $temporary = $path + '.tmp'
    [IO.File]::WriteAllText($temporary, ($value | ConvertTo-Json -Depth 12), $utf8)
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        try {
            if (Test-Path -LiteralPath $path) {
                [IO.File]::Replace($temporary, $path, [NullString]::Value)
            } else {
                [IO.File]::Move($temporary, $path)
            }
            return
        } catch [IO.IOException] {
            if ($attempt -eq 29) { throw }
            Start-Sleep -Milliseconds 10
        }
    }
}

function Set-Stage([string]$name) {
    if ($null -ne $owned) {
        try {
            $process = Get-Process -Id $owned.Pid -ErrorAction Stop
            $result.memory += [ordered]@{
                before_stage = $name; after_stage = $script:stage
                private_bytes = [long]$process.PrivateMemorySize64
                working_set_bytes = [long]$process.WorkingSet64
                peak_working_set_bytes = [long]$process.PeakWorkingSet64
                paged_bytes = [long]$process.PagedMemorySize64
                virtual_bytes = [long]$process.VirtualMemorySize64
            }
        } catch {
            $result.memory += @{ before_stage = $name; available = $false; error = $_.Exception.Message }
        }
    }
    $script:timings[$script:stage] = $script:stageTimer.Elapsed.TotalSeconds
    $script:stage = $name
    $script:stageTimer.Restart()
    Write-Json (Join-Path $directory 'stage.json') @{ stage = $name; updated_utc = [DateTime]::UtcNow.ToString('o') }
}

function Open-Workbook([string]$path) {
    $missing = [Type]::Missing
    if ($config.mode -eq 'verify') {
        # Only a checked, unencrypted XLSX package reaches this phase. Keep the
        # open call minimal: legacy import arguments rejected valid EUSES XLSX.
        return $excel.Workbooks.Open($path, 0, $true)
    }
    # A short explicit nonempty password prevents an omitted-password prompt.
    # It is not a password recovery attempt. Unprotected files ignore it.
    try {
        return $excel.Workbooks.Open($path, 0, $true, $missing, 'Probe-7xQ9!', 'Probe-7xQ9!', $true,
            $missing, $missing, $false, $false, $missing, $false, $true)
    } catch {
        if ($_.Exception.Message -notmatch 'Trust Center|Dateizugriffsschutz|File Block') { throw }
        $result.initial_open_error = $_.Exception.Message
        Set-Stage 'protected_view_open'
        $view = $excel.ProtectedViewWindows.Open($path, 'Probe-7xQ9!', $false, $false)
        try {
            Set-Stage 'enable_editing'
            $book = $view.Edit('Probe-7xQ9!', 0)
            $result.protected_view_edit = $true
            return $book
        } finally {
            if ($null -ne $view) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view) }
        }
    }
}

try {
    Write-Json (Join-Path $directory 'stage.json') @{ stage = 'startup'; updated_utc = [DateTime]::UtcNow.ToString('o') }
    Add-Type -Path (Join-Path $PSScriptRoot 'OwnedExcel.cs')
    $owned = New-Object OwnedExcel($config.excel, $config.seed)
    Write-Json (Join-Path $directory 'process.json') @{
        pid = $owned.Pid; created_filetime = $owned.CreatedFileTime
    }
    # EXCEL7 can exist before its automation Application is ready.
    # The Python supervisor enforces the startup deadline outside these COM calls.
    while ($null -eq $excel) {
        $native = $owned.FindNativeWindow()
        if ($null -ne $native) {
            try {
                $candidate = $native.Application
                if ($null -ne $candidate -and $null -ne $candidate.Hwnd -and [long]$candidate.Hwnd -ne 0) {
                    $attachedPid = [uint32]0
                    [void][OwnedExcel]::GetWindowThreadProcessId([IntPtr]$candidate.Hwnd, [ref]$attachedPid)
                    if ($attachedPid -ne $owned.Pid) { throw 'The Excel window does not belong to this worker.' }
                    $excel = $candidate
                }
            } catch {
                if ($_.Exception.Message -like '*does not belong*') { throw }
            }
        }
        if ($null -eq $excel) { Start-Sleep -Milliseconds 100 }
    }

    Set-Stage 'configure'
    $excel.Visible = $false
    $excel.DisplayAlerts = ($config.mode -eq 'verify')
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false
    $result.excel = [ordered]@{
        version = [string]$excel.Version
        build = [string]$excel.Build
        operating_system = [string]$excel.OperatingSystem
        locale = [Globalization.CultureInfo]::CurrentCulture.Name
        ui_locale = [Globalization.CultureInfo]::CurrentUICulture.Name
    }

    Set-Stage 'open'
    $callTimer = [Diagnostics.Stopwatch]::StartNew()
    $workbook = Open-Workbook $config.source
    $callTimer.Stop()
    $result.open_seconds = $callTimer.Elapsed.TotalSeconds
    if ($null -eq $workbook) { throw 'Excel returned no workbook.' }
    $result.shared_workbook = [bool]$workbook.MultiUserEditing
    $result.sheet_names = @($workbook.Sheets | ForEach-Object { [string]$_.Name })
    $result.sheet_count = [int]$workbook.Sheets.Count
    $result.workbook_structure_protected = [bool]$workbook.ProtectStructure
    if ($config.mode -in @('convert', 'convert_exclusive', 'reconstruct')) {
        $hasVba = [bool]$workbook.HasVBProject
        $macroSheets = 0
        foreach ($sheet in $workbook.Sheets) {
            try { if ([int]$sheet.Type -in @(3, 4)) { $macroSheets++ } }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
        }
        $result.macros_removed = ($hasVba -or ($macroSheets -gt 0))
        $result.vba_project = $hasVba
        $result.xlm_sheet_count = $macroSheets
        if ($config.mode -eq 'reconstruct') {
            Set-Stage 'reconstruction_wait_clipboard'
            $clipboardMutex = New-Object Threading.Mutex($false, 'Local\ExcelShrinkCorpusReconstruction')
            try { $clipboardOwned = $clipboardMutex.WaitOne(60000) }
            catch [Threading.AbandonedMutexException] {
                $clipboardOwned = $true
                $result.previous_reconstruction_owner_abandoned = $true
            }
            if (-not $clipboardOwned) { throw 'Another reconstruction retained the shared clipboard beyond its wait limit.' }
            Set-Stage 'reconstruction_create_sheets'
            # Create every destination title before copying cell contents. Do
            # not silently omit chart sheets, macro sheets, or hidden sheets.
            foreach ($sheet in $workbook.Sheets) {
                if ([int]$sheet.Type -ne -4167) { throw 'Cell-copy reconstruction supports worksheets only; another sheet type requires review.' }
            }
            $rebuilt = $excel.Workbooks.Add(-4167)
            $rebuilt.Date1904 = $workbook.Date1904
            for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
                if ($index -eq 1) { $target = $rebuilt.Worksheets.Item(1) }
                else { $target = $rebuilt.Worksheets.Add([Type]::Missing, $rebuilt.Worksheets.Item($rebuilt.Worksheets.Count)) }
                $target.Name = [string]$workbook.Worksheets.Item($index).Name
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($target)
            }
            $copyObjectsSetting = [bool]$excel.CopyObjectsWithCells
            $excel.CopyObjectsWithCells = $true
            $result.reconstruction = [ordered]@{
                method = 'all_cells_clipboard_paste_into_fresh_workbook'
                equivalent_user_actions = 'Select all cells, copy, paste into a fresh worksheet, retain worksheet names and order'
                fidelity = 'not_established'; sheets = @()
                source_links = @($workbook.LinkSources(1))
                limitations = @('Workbook-level metadata is not cloned', 'Row heights, print settings, names, formulas and embedded objects require review')
            }
            for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
                Set-Stage ('reconstruction_sheet_' + $index)
                $sourceSheet = $workbook.Worksheets.Item($index)
                $targetSheet = $rebuilt.Worksheets.Item($index)
                $sourceCells = $sourceSheet.Cells
                $anchor = $targetSheet.Range('A1')
                try {
                    $beforeShapes = [int]$sourceSheet.Shapes.Count
                    $sourceCells.Copy()
                    $targetSheet.Activate()
                    $targetSheet.Paste($anchor)
                    $result.reconstruction.sheets += [ordered]@{
                        name = [string]$sourceSheet.Name; source_visibility = [int]$sourceSheet.Visible
                        source_shapes = $beforeShapes; output_shapes = [int]$targetSheet.Shapes.Count
                    }
                } finally {
                    $excel.CutCopyMode = $false
                    foreach ($reference in @($anchor, $sourceCells, $targetSheet, $sourceSheet)) {
                        if ($null -ne $reference) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) }
                    }
                }
            }
            # Apply visibility only after all cells have been pasted.
            for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
                $rebuilt.Worksheets.Item($index).Visible = $workbook.Worksheets.Item($index).Visible
            }
            $result.reconstruction.output_links = @($rebuilt.LinkSources(1))
            Set-Stage 'reconstruction_save'
            $rebuilt.CheckCompatibility = $false
            $rebuilt.SaveAs($config.output, 51, '', '', $false, $false)
            $rebuilt.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($rebuilt)
            $rebuilt = $null
        } else {
        Set-Stage 'save'
        $workbook.CheckCompatibility = $false
        $callTimer.Restart()
        if ($config.mode -eq 'convert_exclusive') {
            $workbook.SaveAs($config.output, 51, '', '', $false, $false, 3)
        } else {
            $workbook.SaveAs($config.output, 51, '', '', $false, $false)
        }
        $result.shared_after_save = [bool]$workbook.MultiUserEditing
        $callTimer.Stop()
        $result.save_seconds = $callTimer.Elapsed.TotalSeconds
        }
    }
    Set-Stage 'close'
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    $workbook = $null
    $excel.DisplayAlerts = $false
    $excel.Quit()
    Set-Stage 'complete'
    $result.status = 'accepted'
    $result.reason = 'completed'
} catch {
    $result.reason = 'excel_error'
    $result.error = $_.Exception.Message
    $result.hresult = ('0x{0:X8}' -f ($_.Exception.HResult -band 0xffffffffL))
    $result.error_type = $_.Exception.GetType().FullName
    [Console]::Error.WriteLine($result.error)
} finally {
    if ($null -ne $excel -and $null -ne $copyObjectsSetting) {
        try { $excel.CopyObjectsWithCells = $copyObjectsSetting } catch { }
    }
    $timings[$stage] = $stageTimer.Elapsed.TotalSeconds
    $result.stage = $stage
    Write-Json (Join-Path $directory 'result.json') $result
    # Closing this handle is independent of Excel's COM response. No user's
    # Excel process can belong to this private job.
    if ($null -ne $owned) { $owned.Dispose() }
    foreach ($reference in @($rebuilt, $workbook, $excel, $native)) {
        if ($null -ne $reference) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) } catch { }
        }
    }
    if ($null -ne $clipboardMutex) {
        if ($clipboardOwned) { $clipboardMutex.ReleaseMutex() }
        $clipboardMutex.Dispose()
    }
}
