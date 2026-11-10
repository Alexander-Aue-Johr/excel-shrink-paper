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
$result = [ordered]@{ status = 'review'; stage = 'startup'; timings = $timings }

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
    return $excel.Workbooks.Open($path, 0, $true, $missing, 'Probe-7xQ9!', 'Probe-7xQ9!', $true,
        $missing, $missing, $false, $false, $missing, $false, $true)
}

try {
    Write-Json (Join-Path $directory 'stage.json') @{ stage = 'startup'; updated_utc = [DateTime]::UtcNow.ToString('o') }
    Add-Type -Path (Join-Path $PSScriptRoot 'OwnedExcel.cs')
    $owned = New-Object OwnedExcel($config.excel, $config.seed)
    Write-Json (Join-Path $directory 'process.json') @{
        pid = $owned.Pid; created_filetime = $owned.CreatedFileTime
    }
    while ($null -eq $native) {
        $native = $owned.FindNativeWindow()
        if ($null -eq $native) { Start-Sleep -Milliseconds 100 }
    }
    $excel = $native.Application
    $attachedPid = [uint32]0
    [void][OwnedExcel]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$attachedPid)
    if ($attachedPid -ne $owned.Pid) { throw 'The Excel window does not belong to this worker.' }

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
    if ($config.mode -in @('convert', 'convert_exclusive')) {
        $hasVba = [bool]$workbook.HasVBProject
        $macroSheets = 0
        foreach ($sheet in $workbook.Sheets) {
            try { if ([int]$sheet.Type -in @(3, 4)) { $macroSheets++ } }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
        }
        $result.macros_removed = ($hasVba -or ($macroSheets -gt 0))
        $result.vba_project = $hasVba
        $result.xlm_sheet_count = $macroSheets
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
    $timings[$stage] = $stageTimer.Elapsed.TotalSeconds
    $result.stage = $stage
    Write-Json (Join-Path $directory 'result.json') $result
    # Closing this handle is independent of Excel's COM response. No user's
    # Excel process can belong to this private job.
    if ($null -ne $owned) { $owned.Dispose() }
    foreach ($reference in @($workbook, $excel, $native)) {
        if ($null -ne $reference) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) } catch { }
        }
    }
}
