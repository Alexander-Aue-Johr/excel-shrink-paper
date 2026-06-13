param(
    [int]$WorkbookTimeoutSec = 300,
    [int]$WorkbookTimeoutMaxSec = 600,
    [int]$WorkbookTimeoutAttempts = 3,
    [int]$VisualRows = 60,
    [int]$VisualCols = 20,
    [int]$VisualMaxSheets = 0,
    [int]$VisualPixelTolerance = 0,
    [int]$RustMaxActiveWorkbooks = 1,
    [int]$RustWriterWorkers = 1,
    [int]$MaxVerifyRestarts = 20,
    [int]$RestartDelaySec = 15,
    [int]$VerifyIdleTimeoutSec = 1800,
    [int]$VerifyWatchPollSec = 30,
    [string]$DestatisOut = "",
    [string]$DataGovOut = "",
    [string]$DestatisLog = "",
    [string]$DataGovLog = "",
    [switch]$Force,
    [switch]$NoRecompareAll
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
Set-Location $repoRoot

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$verifyScript = Join-Path $repoRoot "scripts\verify_shrunk_xlsx_values\verify_shrunk_xlsx_values.py"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path $python)) {
    throw "Python venv not found: $python"
}
if (-not (Test-Path $verifyScript)) {
    throw "Verify script not found: $verifyScript"
}

function Invoke-VerifyRun {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Originals,
        [Parameter(Mandatory = $true)][string]$Shrunk,
        [Parameter(Mandatory = $true)][string]$Out,
        [Parameter(Mandatory = $true)][string]$Log,
        [Parameter(Mandatory = $true)][string]$Screenshots,
        [string]$RepairOriginalsDir = ""
    )

    New-Item -ItemType Directory -Force -Path (Split-Path $Out) | Out-Null
    New-Item -ItemType Directory -Force -Path (Split-Path $Log) | Out-Null
    New-Item -ItemType Directory -Force -Path $Screenshots | Out-Null

    $argsList = @(
        $verifyScript,
        $Originals,
        $Shrunk,
        "--out",
        $Out,
        "--log-file",
        $Log,
        "--visual-compare",
        "--visual-rows",
        "$VisualRows",
        "--visual-cols",
        "$VisualCols",
        "--visual-max-sheets",
        "$VisualMaxSheets",
        "--visual-pixel-tolerance",
        "$VisualPixelTolerance",
        "--visual-screenshot-dir",
        $Screenshots,
        "--shrink-before-compare",
        "--rust-max-active-workbooks",
        "$RustMaxActiveWorkbooks",
        "--rust-writer-workers",
        "$RustWriterWorkers",
        "--workbook-timeout-sec",
        "$WorkbookTimeoutSec",
        "--workbook-timeout-max-sec",
        "$WorkbookTimeoutMaxSec",
        "--workbook-timeout-attempts",
        "$WorkbookTimeoutAttempts"
    )

    if ($RepairOriginalsDir) {
        $argsList += @("--repair-originals-dir", $RepairOriginalsDir)
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting verify: $Name"
    Write-Host "  log:    $Log"
    Write-Host "  report: $Out"
    Write-Host "  shots:  $Screenshots"

    $attempt = 0
    while ($true) {
        $attempt += 1
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Verify attempt ${attempt}: $Name"

        $attemptArgs = @($argsList)
        if ($Force -and -not $NoRecompareAll -and $attempt -eq 1) {
            $attemptArgs += "--recompare-all"
        }

        $attemptStart = Get-Date
        $baselinePids = @{}
        Get-Process python,EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
            $baselinePids[$_.Id] = $true
        }

        $proc = Start-Process `
            -FilePath $python `
            -ArgumentList $attemptArgs `
            -WorkingDirectory $repoRoot `
            -WindowStyle Hidden `
            -PassThru

        $lastLogWrite = $attemptStart
        if (Test-Path $Log) {
            $lastLogWrite = (Get-Item $Log).LastWriteTime
        }

        $timedOut = $false
        while (-not $proc.HasExited) {
            Start-Sleep -Seconds $VerifyWatchPollSec
            if (Test-Path $Log) {
                $currentLogWrite = (Get-Item $Log).LastWriteTime
                if ($currentLogWrite -gt $lastLogWrite) {
                    $lastLogWrite = $currentLogWrite
                }
            }

            $idleSec = ((Get-Date) - $lastLogWrite).TotalSeconds
            if ($idleSec -ge $VerifyIdleTimeoutSec) {
                $timedOut = $true
                Write-Warning "Verify attempt ${attempt} for '$Name' produced no log output for $([int]$idleSec)s. Killing Python/Excel processes started since $attemptStart and resuming."

                Get-Process python,EXCEL -ErrorAction SilentlyContinue | Where-Object {
                    -not $baselinePids.ContainsKey($_.Id) -and $_.StartTime -ge $attemptStart
                } | ForEach-Object {
                    try {
                        Stop-Process -Id $_.Id -Force
                        Write-Warning "Stopped $($_.ProcessName) pid=$($_.Id) started=$($_.StartTime)"
                    } catch {
                        Write-Warning "Could not stop $($_.ProcessName) pid=$($_.Id): $($_.Exception.Message)"
                    }
                }
                break
            }
        }

        if ($timedOut) {
            $exitCode = -999
        } else {
            $proc.WaitForExit()
            $exitCode = $proc.ExitCode
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished verify attempt ${attempt}: $Name with exit code $exitCode"

        if ($exitCode -eq 0) {
            return
        }

        if ($attempt -gt $MaxVerifyRestarts) {
            throw "Verify failed for $Name with exit code $exitCode after $attempt attempts"
        }

        Write-Warning "Verify process for '$Name' exited with code $exitCode. Restarting in $RestartDelaySec seconds; completed workbooks will be resumed from the progress journal next to the report."
        Start-Sleep -Seconds $RestartDelaySec
    }
}

$destatisShrunk = "scripts\analyze_destatis_xlsx_files\shrinked_by_excel_shrink_rust"
if (-not $DestatisOut) {
    $DestatisOut = "scripts\analyze_destatis_xlsx_files\report_rust_visual_$stamp.json"
}
if (-not $DestatisLog) {
    $DestatisLog = "scripts\analyze_destatis_xlsx_files\verify_destatis_rust_visual_$stamp.log"
}
$destatisScreenshots = "scripts\analyze_destatis_xlsx_files\visual_screenshots_$stamp"

$dataGovShrunk = "scripts\analyze_data_gov_xlsx_files\shrinked_by_excel_shrink_rust"
if (-not $DataGovOut) {
    $DataGovOut = "scripts\analyze_data_gov_xlsx_files\report_rust_visual_$stamp.json"
}
if (-not $DataGovLog) {
    $DataGovLog = "scripts\analyze_data_gov_xlsx_files\verify_data_gov_rust_visual_$stamp.log"
}
$dataGovScreenshots = "scripts\analyze_data_gov_xlsx_files\visual_screenshots_$stamp"

Invoke-VerifyRun `
    -Name "Destatis rust shrink-before-verify with visual compare" `
    -Originals "scripts\scrape_destatis_xlsx_files\downloaded" `
    -Shrunk $destatisShrunk `
    -Out $DestatisOut `
    -Log $DestatisLog `
    -Screenshots $destatisScreenshots `
    -RepairOriginalsDir "scripts\analyze_destatis_xlsx_files\fixed_original_corrupt_destatis_files"

Invoke-VerifyRun `
    -Name "data.gov rust shrink-before-verify with visual compare" `
    -Originals "scripts\scrape_data_gov_xlsx_files\downloaded" `
    -Shrunk $dataGovShrunk `
    -Out $DataGovOut `
    -Log $DataGovLog `
    -Screenshots $dataGovScreenshots

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] All shrink-before-verify runs finished."
